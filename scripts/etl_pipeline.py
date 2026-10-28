
import os
import logging
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, URL, text

# --------------------------------------------------
# 1. Project configuration
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_FILE = BASE_DIR / "data" / "raw" / "US_Accidents_March23.csv"
LOG_DIR = BASE_DIR / "logs"

LOG_DIR.mkdir(parents=True, exist_ok=True)

load_dotenv(BASE_DIR / ".env")

CHUNK_SIZE = 10_000
INSERT_BATCH_SIZE = 500
TABLE_NAME = "traffic_accidents"

# --------------------------------------------------
# 2. Logging configuration
# --------------------------------------------------

logging.basicConfig(
    filename=LOG_DIR / "etl_pipeline.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger().addHandler(console)

# --------------------------------------------------
# 3. PostgreSQL connection
# --------------------------------------------------

connection_url = URL.create(
    drivername="postgresql+psycopg",
    username=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD"),
    host=os.getenv("DB_HOST"),
    port=int(os.getenv("DB_PORT", "5432")),
    database=os.getenv("DB_NAME"),
)

engine = create_engine(connection_url, pool_pre_ping=True)

# CSV column names mapped to PostgreSQL column names.
COLUMN_MAP = {
    "ID": "id",
    "Source": "source",
    "Severity": "severity",
    "Start_Time": "start_time",
    "End_Time": "end_time",
    "Start_Lat": "start_lat",
    "Start_Lng": "start_lng",
    "End_Lat": "end_lat",
    "End_Lng": "end_lng",
    "Distance(mi)": "distance_mi",
    "Description": "description",
    "Street": "street",
    "City": "city",
    "County": "county",
    "State": "state",
    "Zipcode": "zipcode",
    "Country": "country",
    "Timezone": "timezone",
    "Airport_Code": "airport_code",
    "Weather_Timestamp": "weather_timestamp",
    "Temperature(F)": "temperature_f",
    "Wind_Chill(F)": "wind_chill_f",
    "Humidity(%)": "humidity_pct",
    "Pressure(in)": "pressure_in",
    "Visibility(mi)": "visibility_mi",
    "Wind_Direction": "wind_direction",
    "Wind_Speed(mph)": "wind_speed_mph",
    "Precipitation(in)": "precipitation_in",
    "Weather_Condition": "weather_condition",
    "Amenity": "amenity",
    "Bump": "bump",
    "Crossing": "crossing",
    "Give_Way": "give_way",
    "Junction": "junction",
    "No_Exit": "no_exit",
    "Railway": "railway",
    "Roundabout": "roundabout",
    "Station": "station",
    "Stop": "stop",
    "Traffic_Calming": "traffic_calming",
    "Traffic_Signal": "traffic_signal",
    "Turning_Loop": "turning_loop",
    "Sunrise_Sunset": "sunrise_sunset",
    "Civil_Twilight": "civil_twilight",
    "Nautical_Twilight": "nautical_twilight",
    "Astronomical_Twilight": "astronomical_twilight",
}

TIMESTAMP_COLUMNS = [
    "Start_Time",
    "End_Time",
    "Weather_Timestamp",
]

NUMERIC_COLUMNS = [
    "Severity",
    "Start_Lat",
    "Start_Lng",
    "End_Lat",
    "End_Lng",
    "Distance(mi)",
    "Temperature(F)",
    "Wind_Chill(F)",
    "Humidity(%)",
    "Pressure(in)",
    "Visibility(mi)",
    "Wind_Speed(mph)",
    "Precipitation(in)",
]

BOOLEAN_COLUMNS = [
    "Amenity",
    "Bump",
    "Crossing",
    "Give_Way",
    "Junction",
    "No_Exit",
    "Railway",
    "Roundabout",
    "Station",
    "Stop",
    "Traffic_Calming",
    "Traffic_Signal",
    "Turning_Loop",
]

# --------------------------------------------------
# 4. Extract and transform
# --------------------------------------------------

def transform_chunk(df):
    """Clean and transform one chunk of accident records."""

    df = df.rename(columns=lambda name: name.strip())

    missing_columns = set(COLUMN_MAP) - set(df.columns)

    if missing_columns:
        raise ValueError(
            f"CSV is missing expected columns: {sorted(missing_columns)}"
        )

    # Keep only the columns used by our database.
    df = df[list(COLUMN_MAP)].copy()

    # Remove rows without an accident ID.
    df["ID"] = df["ID"].astype("string").str.strip()
    df = df[df["ID"].notna() & df["ID"].ne("")]

    # Convert timestamps.
    for column in TIMESTAMP_COLUMNS:
        df[column] = pd.to_datetime(
            df[column],
            errors="coerce",
        )

    # Convert numeric values.
    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    # Convert Boolean columns.
    for column in BOOLEAN_COLUMNS:
        df[column] = df[column].map(
            {
                True: True,
                False: False,
                "True": True,
                "False": False,
                "true": True,
                "false": False,
                1: True,
                0: False,
            }
        ).astype("boolean")

    # Clean text and convert empty strings to missing values.
    for column in df.select_dtypes(
        include=["object", "string"]
    ).columns:
        df[column] = df[column].str.strip()
        df[column] = df[column].replace("", pd.NA)

    # Keep valid accident severity levels.
    df = df[df["Severity"].isin([1, 2, 3, 4])]

    # Remove duplicate IDs within this chunk.
    df = df.drop_duplicates(subset=["ID"])

    # Rename columns to match PostgreSQL.
    return df.rename(columns=COLUMN_MAP)


# --------------------------------------------------
# 5. Load into PostgreSQL
# --------------------------------------------------

def main():
    if not DATA_FILE.is_file():
        raise FileNotFoundError(
            f"Dataset not found: {DATA_FILE}"
        )

    logging.info("Starting traffic accident ETL pipeline.")
    logging.info("Source file: %s", DATA_FILE)
    logging.info("Read chunk size: %s", CHUNK_SIZE)
    logging.info("Insert batch size: %s", INSERT_BATCH_SIZE)

    total_read = 0
    total_loaded = 0
    total_excluded = 0
    chunk_number = 0

    try:
        # Check the database connection and target table.
        with engine.connect() as connection:
            database_name = connection.execute(
                text("SELECT current_database()")
            ).scalar()

            existing_rows = connection.execute(
                text("SELECT COUNT(*) FROM traffic_accidents")
            ).scalar()

            logging.info(
                "Connected to database: %s",
                database_name,
            )

            logging.info(
                "Existing rows in target table: %s",
                existing_rows,
            )

            # Protect against accidental duplicate loads.
            if existing_rows:
                raise RuntimeError(
                    f"The table {TABLE_NAME} already contains "
                    f"{existing_rows} rows. This initial-load pipeline "
                    "requires an empty table. Verify the existing data "
                    "before deciding whether to clear the table."
                )

        # Read the large CSV in chunks rather than loading it all at once.
        for chunk in pd.read_csv(
            DATA_FILE,
            chunksize=CHUNK_SIZE,
            low_memory=False,
        ):
            chunk_number += 1
            rows_read = len(chunk)
            total_read += rows_read

            cleaned = transform_chunk(chunk)
            rows_loaded_this_chunk = len(cleaned)
            total_excluded += rows_read - rows_loaded_this_chunk

            if not cleaned.empty:
                # Commit each chunk as a separate transaction.
                with engine.begin() as connection:
                    cleaned.to_sql(
                        TABLE_NAME,
                        con=connection,
                        if_exists="append",
                        index=False,
                        chunksize=INSERT_BATCH_SIZE,
                    )

                total_loaded += rows_loaded_this_chunk

            logging.info(
                "Chunk %s: read=%s, loaded=%s, excluded=%s, "
                "total_read=%s, total_loaded=%s",
                chunk_number,
                rows_read,
                rows_loaded_this_chunk,
                rows_read - rows_loaded_this_chunk,
                total_read,
                total_loaded,
            )

        logging.info("ETL completed successfully.")
        logging.info("Total rows read: %s", total_read)
        logging.info("Total rows loaded: %s", total_loaded)
        logging.info("Total rows excluded: %s", total_excluded)

        print("\nETL COMPLETED SUCCESSFULLY")
        print("Rows read:", total_read)
        print("Rows loaded:", total_loaded)
        print("Rows excluded during transformation:", total_excluded)

    except Exception:
        logging.exception("ETL pipeline failed.")
        raise

    finally:
        engine.dispose()


if __name__ == "__main__":
    main()

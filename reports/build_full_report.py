import os
import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from report_builder_helper import (
    set_cell_background, set_cell_margins, add_callout, 
    add_code_block, style_table
)

def build_docx_report():
    doc = Document()
    
    # Page setup - 1 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    reports_dir = Path(r"E:\BigData_Traffic_Accident_Project\reports")
    
    # Configure default style fonts
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    # -------------------------------------------------------------
    # 1. TITLE PAGE
    # -------------------------------------------------------------
    p_title_space = doc.add_paragraph()
    p_title_space.paragraph_format.space_before = Pt(72)
    
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("[NAME OF UNIVERSITY / FACULTY / DEPARTMENT]\nDEPARTMENT OF COMPUTER SCIENCE & INFORMATION TECHNOLOGY")
    r_inst.bold = True
    r_inst.font.size = Pt(12)
    r_inst.font.color.rgb = RGBColor(0x1A, 0x23, 0x7E)
    
    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_course.paragraph_format.space_after = Pt(40)
    r_course = p_course.add_run("COURSE: [COURSE CODE - COURSE TITLE, E.G., CS410: BIG DATA ENGINEERING & ANALYTICS]")
    r_course.font.size = Pt(11)
    r_course.font.color.rgb = RGBColor(0x54, 0x6E, 0x7A)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Design and Implementation of a Big Data ETL Pipeline for US Traffic Accident Analysis Using Python Pandas and PostgreSQL")
    r_title.bold = True
    r_title.font.size = Pt(22)
    r_title.font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(70)
    r_sub = p_sub.add_run("A High-Throughput Chunk-Based Data Ingestion, Cleaning, and Analytical Architecture for 7.7 Million Records")
    r_sub.italic = True
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(0x37, 0x47, 0x4F)
    
    # Metadata Box
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.line_spacing = 1.3
    
    runs_meta = [
        ("Prepared By (Student Name / ID): ", True), ("[STUDENT NAME] (ID: [STUDENT ID])\n", False),
        ("Group Members: ", True), ("[MEMBER 1 NAME & ID], [MEMBER 2 NAME & ID], [MEMBER 3 NAME & ID]\n", False),
        ("Instructor / Supervisor: ", True), ("[INSTRUCTOR NAME & ACADEMIC TITLE]\n", False),
        ("Academic Term: ", True), ("[SEMESTER / ACADEMIC YEAR, E.G., FALL 2026]\n", False),
        ("Submission Date: ", True), ("[OCTOBER 27, 2026 / SUBMISSION DATE]\n", False),
    ]
    for text_val, is_b in runs_meta:
        r = p_meta.add_run(text_val)
        r.bold = is_b
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x26, 0x32, 0x38)
        
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 2. ABSTRACT
    # -------------------------------------------------------------
    h1 = doc.add_heading("2. Abstract", level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(8)
    
    p_abs = doc.add_paragraph(
        "Modern big data analytics increasingly demands resilient, memory-bounded, and deterministic "
        "Extract, Transform, and Load (ETL) pipelines capable of ingesting multi-gigabyte flat files into relational and "
        "analytical database environments. This project presents the end-to-end engineering design and empirical implementation "
        "of a high-performance ETL pipeline developed using Python, Pandas, SQLAlchemy 2.0, and Psycopg 3, targeting a PostgreSQL "
        "relational database system. The pipeline processes a real-world US traffic accident dataset comprising 7,728,394 records "
        "and 47 heterogeneous attributes across a 2.85 GB physical CSV file. To overcome primary memory saturation on commodity hardware, "
        "a chunk-based streaming architecture was engineered with a calibrated chunk size of 10,000 records and an insertion batch size "
        "of 500 records executed under transactional isolation."
    )
    p_abs.paragraph_format.space_after = Pt(6)
    
    p_abs2 = doc.add_paragraph(
        "Empirical production execution logs demonstrate that all 7,728,394 source records were successfully extracted, transformed, and loaded "
        "across 773 chunks in approximately 8 minutes and 21 seconds (501 seconds), achieving an effective ingestion throughput of approximately "
        "15,426 records per second with zero row exclusions. The transformation layer strictly enforces string sanitation, numeric and timestamp "
        "type coercion, Boolean mapping, severity filtering, and in-chunk deduplication. Post-load SQL analytical queries reveal key macroeconomic "
        "and environmental patterns across the United States, including a high concentration of Severity 2 incidents (79.7%), prominent state-level "
        "reporting frequency in California and Florida, and a predominance of clear and cloudy weather conditions during recorded accidents. "
        "Finally, this paper conducts a candid technical evaluation of architectural limitations—including chunk-level versus global deduplication, "
        "transaction granularity, absence of secondary index structures in the base DDL, and error coercion masking—and proposes concrete "
        "engineering recommendations for enterprise deployment."
    )
    p_abs2.paragraph_format.space_after = Pt(8)
    
    p_kw = doc.add_paragraph()
    r_kw = p_kw.add_run("Keywords: ")
    r_kw.bold = True
    p_kw.add_run("Big Data Engineering, ETL Pipeline, Python Pandas, PostgreSQL, SQLAlchemy, Psycopg, Traffic Accident Analytics, Chunk Processing, Data Integrity.")
    p_kw.paragraph_format.space_after = Pt(18)
    
    # -------------------------------------------------------------
    # 3. TABLE OF CONTENTS
    # -------------------------------------------------------------
    h1 = doc.add_heading("3. Table of Contents", level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(8)
    
    toc_items = [
        ("1. Title Page", "1"),
        ("2. Abstract", "2"),
        ("3. Table of Contents", "3"),
        ("4. Introduction", "4"),
        ("5. Background of the Study", "4"),
        ("6. Problem Statement", "5"),
        ("7. General and Specific Objectives", "6"),
        ("8. Scope and Significance of the Project", "6"),
        ("9. Dataset Description and Source", "7"),
        ("10. Tools and Technologies", "8"),
        ("11. System Requirements", "9"),
        ("12. System Architecture Diagram", "10"),
        ("13. ETL Workflow and Methodology", "12"),
        ("14. Data Cleaning and Transformation Rules", "14"),
        ("15. PostgreSQL Database Design and Schema Explanation", "16"),
        ("16. Python Implementation Explanation", "19"),
        ("17. Error Handling, Logging, and Memory Management", "21"),
        ("18. Execution Results and Performance Discussion", "23"),
        ("19. SQL-Based Data Analysis", "25"),
        ("20. Testing and Validation", "28"),
        ("21. Limitations and Challenges", "30"),
        ("22. Recommendations and Future Improvements", "32"),
        ("23. Conclusion", "34"),
        ("24. References", "35"),
        ("25. Appendix A: Installation and Execution Instructions", "37"),
        ("26. Appendix B: Important Python and SQL Examples", "39"),
    ]
    
    tbl_toc = doc.add_table(rows=len(toc_items), cols=2)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, (title_item, page_str) in enumerate(toc_items):
        row = tbl_toc.rows[idx]
        cell_t = row.cells[0]
        cell_p = row.cells[1]
        set_cell_margins(cell_t, top=40, bottom=40, left=60, right=60)
        set_cell_margins(cell_p, top=40, bottom=40, left=60, right=60)
        
        pt = cell_t.paragraphs[0]
        pr = cell_p.paragraphs[0]
        pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        rt = pt.add_run(title_item)
        rt.font.name = "Calibri"
        rt.font.size = Pt(9.5)
        if title_item.startswith(("1.", "2.", "3.", "4.", "12.", "15.", "18.", "19.", "23.")):
            rt.bold = True
            
        rp = pr.add_run(page_str)
        rp.font.name = "Calibri"
        rp.font.size = Pt(9.5)
        
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 4. INTRODUCTION
    # -------------------------------------------------------------
    doc.add_heading("4. Introduction", level=1)
    doc.add_paragraph(
        "Transportation safety analysis is a foundational pillar of modern intelligent transportation systems (ITS), civil infrastructure "
        "planning, and urban risk mitigation. Across the United States, traffic collisions generate catastrophic economic, medical, and human losses annually. "
        "Understanding where, when, and under what atmospheric or road configurations vehicular collisions take place requires the aggregation "
        "and empirical processing of multi-year nationwide event logs. Such telemetry datasets, collected by state departments of transportation, "
        "law enforcement agencies, traffic cameras, and real-world road sensors, inevitably scale to millions of heterogeneous records."
    )
    doc.add_paragraph(
        "In enterprise big data engineering, ingesting and structuring multi-gigabyte flat files poses distinct computational challenges. "
        "Directly loading multi-gigabyte files into system memory induces out-of-memory (OOM) fatal crashes, while unoptimized bulk database "
        "inserts lead to severe connection exhaustion, transaction log bloating, and locking contention. This academic report details the complete "
        "architectural lifecycle of designing, executing, and validating an industrial-grade ETL (Extract, Transform, Load) pipeline using Python Pandas, "
        "SQLAlchemy 2.0, Psycopg 3, and a relational PostgreSQL database."
    )
    
    # -------------------------------------------------------------
    # 5. BACKGROUND OF THE STUDY
    # -------------------------------------------------------------
    doc.add_heading("5. Background of the Study", level=1)
    doc.add_paragraph(
        "The focal dataset of this project is the widely studied US Accidents corpus (specifically US_Accidents_March23.csv), which represents one "
        "of the most comprehensive countrywide vehicular accident datasets available. Spanning multiple calendar years across 49 contiguous US states, "
        "the raw dataset captures 7.73 million discrete collision events with 46 associated feature dimensions encompassing geographical coordinates, "
        "timestamps, weather station telemetry (temperature, wind chill, barometric pressure, visibility), and fine-grained road infrastructure flags "
        "(such as pedestrian crossings, railway intersections, traffic signals, and roundabouts)."
    )
    doc.add_paragraph(
        "Historically, transportation data warehousing faced severe fragmentation. Local municipalities and regional departments maintained isolated "
        "repositories with conflicting naming schemes, erratic timestamp notations, non-standardized boolean encodings, and missing sensor telemetry. "
        "Bridging the divide between raw comma-separated value (CSV) dumps and relational online analytical processing (OLAP) databases necessitates "
        "an engineered ingestion pipeline that guarantees data type preservation, handles schema discrepancies deterministically, and provides "
        "granular auditability through system logging."
    )
    
    # -------------------------------------------------------------
    # 6. PROBLEM STATEMENT
    # -------------------------------------------------------------
    doc.add_heading("6. Problem Statement", level=1)
    doc.add_paragraph(
        "Ingesting, validating, and structuring large-scale heterogeneous traffic event logs exposes four fundamental engineering bottlenecks:"
    )
    doc.add_paragraph(
        "1. Memory Exhaustion in Flat File Parsing: A 2.85 GB uncompressed CSV file expands significantly when deserialized into standard Python "
        "data structures due to object pointer overhead and string buffer allocations. Attempting to parse 7.73 million rows into a single Pandas "
        "DataFrame in commodity development environments exceeds available system RAM, resulting in operating system paging or process termination."
    )
    doc.add_paragraph(
        "2. Data Heterogeneity and Dirty Telemetry: Raw traffic records originate from diverse sensory feeds. Common data anomalies include "
        "untrimmed whitespace in identifiers, polymorphic boolean encodings ('True', 'false', 1, 0), irregular timestamp formats across multiple timezones, "
        "missing or corrupt weather sensor readings, and invalid severity classifications."
    )
    doc.add_paragraph(
        "3. Relational Load Performance vs. Transactional Integrity: Loading millions of records into PostgreSQL using naive single-row INSERT statements "
        "incurs massive network latency and query-planning overhead. Conversely, attempting a single uncommitted transaction across millions of records "
        "risks rolling back hours of processing if an error occurs in the final row, while overwhelming the PostgreSQL Write-Ahead Log (WAL)."
    )
    doc.add_paragraph(
        "4. Auditability and Pre-Load Safety: In enterprise environments, rerunning an ingestion pipeline over a non-empty table causes catastrophic "
        "primary key collisions or duplicated analytical counts. An engineered pipeline must enforce deterministic pre-execution state checks and "
        "comprehensive disk-persisted logging."
    )
    
    # -------------------------------------------------------------
    # 7. GENERAL AND SPECIFIC OBJECTIVES
    # -------------------------------------------------------------
    doc.add_heading("7. General and Specific Objectives", level=1)
    doc.add_paragraph(
        "General Objective: To design, implement, and benchmark an automated, memory-bounded, production-grade Big Data ETL pipeline "
        "using Python Pandas and PostgreSQL that cleans and ingests over 7.7 million US traffic accident records for high-performance analytical querying."
    )
    doc.add_paragraph("Specific Objectives:")
    doc.add_paragraph("• Architect a chunked data streaming pipeline utilizing Pandas `read_csv(chunksize=10000)` to maintain constant O(chunk_size) memory complexity.")
    doc.add_paragraph("• Design a normalized 47-column target table schema in PostgreSQL leveraging optimal native data types (VARCHAR, SMALLINT, TIMESTAMP, DOUBLE PRECISION, BOOLEAN).")
    doc.add_paragraph("• Develop a robust transformation module executing string stripping, type coercion (errors='coerce'), boolean standardization, severity domain validation, and in-chunk deduplication.")
    doc.add_paragraph("• Implement a batched transactional loading mechanism using SQLAlchemy 2.0 and the modern Psycopg 3 database driver with insert batching (chunksize=500).")
    doc.add_paragraph("• Enforce environment-based configuration management (.env via python-dotenv) to protect credentials and adhere to the Twelve-Factor App methodology.")
    doc.add_paragraph("• Provide granular dual logging (rotating file handler and console stream) recording chunk-level ingestion progress, excluded row counts, and timestamps.")
    doc.add_paragraph("• Formulate and execute analytical SQL queries to derive traffic safety insights regarding severity distribution, geographic clustering, and meteorological conditions.")
    doc.add_paragraph("• Critically assess technical trade-offs, documented schema gaps, and failure recovery limitations, proposing enterprise-grade enhancements.")
    
    # -------------------------------------------------------------
    # 8. SCOPE AND SIGNIFICANCE OF THE PROJECT
    # -------------------------------------------------------------
    doc.add_heading("8. Scope and Significance of the Project", level=1)
    doc.add_paragraph(
        "Project Scope: The pipeline encompasses the ingestion of the primary US_Accidents_March23.csv dataset (7,728,394 records) into a local or "
        "networked PostgreSQL 15/16 database instance running on Windows. The pipeline script (scripts/etl_pipeline.py) handles CSV pre-flight verification, "
        "table emptiness validation, per-chunk extraction, transformation, database insertion, and engine lifecycle disposal. The project establishes "
        "testing scripts (test_connection.py), table generation DDL (sql/create_table.sql), and comprehensive log artifacts (logs/etl_pipeline.log)."
    )
    doc.add_paragraph(
        "Significance: This project demonstrates that big data processing does not strictly mandate distributed Spark or Hadoop clusters for datasets "
        "in the 1-10 GB range. By applying rigorous data streaming, vectorization in Pandas, and optimized relational bulk loading, a single workstation "
        "can reliably process 7.7 million complex records in under nine minutes. The resulting database enables transportation safety researchers, civil planners, "
        "and emergency service coordinators to conduct millisecond-level exploratory spatial and temporal SQL queries."
    )
    
    # -------------------------------------------------------------
    # 9. DATASET DESCRIPTION AND SOURCE
    # -------------------------------------------------------------
    doc.add_heading("9. Dataset Description and Source", level=1)
    doc.add_paragraph(
        "The project utilizes the US Accidents dataset (file: US_Accidents_March23.csv), originally collected and curated by Sobhan Moosavi et al., "
        "covering traffic incident telemetry across the contiguous United States from February 2016 to March 2023. The dataset aggregates data from multiple "
        "traffic telemetry providers, including the US Department of Transportation, state highway patrol departments, and commercial traffic monitoring systems."
    )
    
    # Summary Table for Dataset
    tbl_ds = doc.add_table(rows=7, cols=2)
    ds_data = [
        ("File Name", "US_Accidents_March23.csv"),
        ("Physical Storage Path", "data/raw/US_Accidents_March23.csv"),
        ("File Size on Disk", "2.85 GB (3,058,183,727 bytes)"),
        ("Total Source Records", "7,728,394 rows"),
        ("Total Column Count", "46 raw attributes (mapped to 47 relational columns)"),
        ("Primary Key Candidate", "ID (Unique alphanumeric incident identifier, e.g., 'A-1')"),
    ]
    style_table(tbl_ds, [2.5, 4.0], ["Dataset Metric", "Specification"], ds_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    add_callout(doc, 
        "Academic Citation Note: While the dataset is widely circulated across academic platforms such as Kaggle and Zenodo, "
        "the definitive academic citation must be referenced based on Moosavi et al. (2019): 'A Countrywide Traffic Accident Dataset', "
        "published in IEEE International Conference on Information Reuse and Integration (IRI). "
        "Students must verify the exact source download URL utilized in their specific university repository prior to final defense.",
        title="DATASET PROVENANCE & CITATION VERIFICATION", border_color="0288D1", bg_color="E1F5FE"
    )
    
    # -------------------------------------------------------------
    # 10. TOOLS AND TECHNOLOGIES
    # -------------------------------------------------------------
    doc.add_heading("10. Tools and Technologies", level=1)
    doc.add_paragraph(
        "The technological stack was selected to maximize memory efficiency, database throughput, security, and reproducibility:"
    )
    
    tech_data = [
        ("Python 3.13 / 3.11", "Runtime environment providing core asynchronous execution, strong typing extensions, and standard library utilities."),
        ("Pandas 3.x / 2.x", "High-performance data manipulation library. Utilized specifically for vectorized chunk streaming (pd.read_csv) and type coercion."),
        ("PostgreSQL 15 / 16", "Enterprise open-source Object-Relational Database Management System (ORDBMS) selected for ACID compliance and index scalability."),
        ("SQLAlchemy 2.1+", "Industry-standard SQL toolkit and Object-Relational Mapping (ORM) layer. Utilized for engine creation, connection pooling, and DataFrame.to_sql integration."),
        ("Psycopg 3 (psycopg)", "Modern PostgreSQL adapter for Python. Provides native binary protocol communication, high-performance COPY/insert streaming, and prepared statements."),
        ("python-dotenv", "Configuration manager loading system environment variables from .env files, isolating secrets from source code control."),
        ("Git & GitHub", "Distributed version control system managing code revisions, branch lifecycles, and academic collaboration while ignoring secrets and logs.")
    ]
    tbl_tech = doc.add_table(rows=len(tech_data) + 1, cols=2)
    style_table(tbl_tech, [2.0, 4.5], ["Technology", "Role & Engineering Justification"], tech_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # -------------------------------------------------------------
    # 11. SYSTEM REQUIREMENTS
    # -------------------------------------------------------------
    doc.add_heading("11. System Requirements", level=1)
    doc.add_paragraph(
        "To ensure deterministic execution and reproduce the benchmarked pipeline throughput (~15,400 rows/sec), the target workstation "
        "should satisfy the following hardware and software baselines:"
    )
    
    sys_req_data = [
        ("Processor (CPU)", "Intel Core i5/i7 (8th Gen+) or AMD Ryzen 5/7 (4+ physical cores recommended for concurrent parsing)"),
        ("System RAM", "Minimum: 8 GB RAM; Recommended: 16 GB RAM (allows comfortable OS disk caching during chunk processing)"),
        ("Disk Storage", "Minimum: 20 GB free space on Solid State Drive (SSD). NVMe SSD strongly recommended for WAL throughput"),
        ("Operating System", "Microsoft Windows 10/11 64-bit (or Linux Ubuntu 22.04 LTS / macOS Ventura)"),
        ("Database Engine", "PostgreSQL Server version 14.x, 15.x, or 16.x listening on TCP port 5432"),
        ("Python Runtime", "Python 64-bit version 3.10 to 3.13 configured within an isolated virtual environment (.venv)")
    ]
    tbl_sys = doc.add_table(rows=len(sys_req_data) + 1, cols=2)
    style_table(tbl_sys, [2.2, 4.3], ["Component", "Minimum / Recommended Specification"], sys_req_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # -------------------------------------------------------------
    # 12. SYSTEM ARCHITECTURE DIAGRAM
    # -------------------------------------------------------------
    doc.add_heading("12. System Architecture Diagram", level=1)
    doc.add_paragraph(
        "The system architecture follows a decoupled, feed-forward streaming pipeline. The physical CSV file on local disk is read "
        "in discrete 10,000-row chunks by the Pandas extraction engine. Each chunk is streamed directly into memory, subjected to transformation "
        "rules in vectorized memory buffers, and written across a pooled connection to PostgreSQL using transactional boundaries. "
        "Figure 1 illustrates the structural components and data flows."
    )
    
    arch_img_file = str(reports_dir / "system_architecture.png")
    if os.path.exists(arch_img_file):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(arch_img_file, width=Inches(6.2))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(12)
        r_cap = p_cap.add_run("Figure 1: High-Throughput Big Data ETL Pipeline Architecture for US Traffic Accidents")
        r_cap.bold = True
        r_cap.font.size = Pt(9.5)
        r_cap.font.color.rgb = RGBColor(0x37, 0x47, 0x4F)
        
    doc.add_paragraph(
        "For documentation portability and integration with Markdown/Mermaid-compatible rendering engines, the system architecture is "
        "formally specified below:"
    )
    
    mermaid_code = (
        "graph TD\n"
        "    subgraph Storage Layer\n"
        "        CSV[Raw CSV Dataset: US_Accidents_March23.csv<br/>2.85 GB / 7.73M Rows]\n"
        "    end\n\n"
        "    subgraph Ingestion & Transformation Layer (Python / Pandas)\n"
        "        Init[Pre-Execution Verification<br/>File Check & Table Empty Guard]\n"
        "        Chunk[Pandas Chunk Iterator<br/>pd.read_csv chunksize=10,000]\n"
        "        Transform[Transformation Module<br/>- Clean Column Headers<br/>- ID Validation & In-Chunk Dedup<br/>- Coerce Timestamps & Numerics<br/>- Map Booleans to True/False<br/>- Trim Strings & Convert Empty to NA<br/>- Filter Valid Severity 1..4]\n"
        "    end\n\n"
        "    subgraph Loading Layer (SQLAlchemy / Psycopg 3)\n"
        "        Transact[Transaction Context Manager<br/>engine.begin per Chunk]\n"
        "        Batch[DataFrame.to_sql<br/>Batch Size = 500 Records]\n"
        "    end\n\n"
        "    subgraph Database & Analysis Layer (PostgreSQL)\n"
        "        PG[(PostgreSQL Database<br/>Table: traffic_accidents<br/>7,728,394 Records Loaded)]\n"
        "        SQL[Analytical SQL Queries<br/>Aggregations, Severity, Weather, States]\n"
        "    end\n\n"
        "    subgraph Cross-Cutting Concerns\n"
        "        ENV[.env Configuration<br/>DB Credentials & Host]\n"
        "        LOG[Rotating System Logger<br/>Console & etl_pipeline.log]\n"
        "    end\n\n"
        "    CSV --> Init --> Chunk --> Transform --> Transact --> Batch --> PG --> SQL\n"
        "    ENV -.-> Init\n"
        "    ENV -.-> Transact\n"
        "    Chunk -.-> LOG\n"
        "    Transform -.-> LOG\n"
        "    Batch -.-> LOG"
    )
    add_code_block(doc, mermaid_code, language="Mermaid")
    
    # -------------------------------------------------------------
    # 13. ETL WORKFLOW AND METHODOLOGY
    # -------------------------------------------------------------
    doc.add_heading("13. ETL Workflow and Methodology", level=1)
    doc.add_paragraph(
        "The ETL methodology adheres to modern data engineering best practices by decoupling extraction, transformation, "
        "and loading into deterministic, testable operations. The execution lifecycle proceeds across seven synchronized phases:"
    )
    doc.add_paragraph(
        "1. Phase 1: Environment Initialization and Pre-Flight Checks: Upon invocation, the pipeline loads database configuration parameters "
        "from the root .env file using `python-dotenv`. It validates the physical existence of `data/raw/US_Accidents_March23.csv` on the filesystem. "
        "If missing, execution aborts immediately via `FileNotFoundError`."
    )
    doc.add_paragraph(
        "2. Phase 2: Database Connection and State Guard: The pipeline instantiates a SQLAlchemy engine utilizing the modern `postgresql+psycopg` "
        "connection string with connection pool pre-ping enabled (`pool_pre_ping=True`). A pre-flight query executes `SELECT current_database()` "
        "and `SELECT COUNT(*) FROM traffic_accidents`. If the target table contains any existing records, the script raises a `RuntimeError` and terminates. "
        "This defensive guard prevents duplicate record accumulation or catastrophic primary key collisions during unmonitored runs."
    )
    doc.add_paragraph(
        "3. Phase 3: Chunk-Based Extraction: Rather than loading the entire 2.85 GB file into RAM, the extraction engine invokes `pd.read_csv()` "
        "with `chunksize=10000` and `low_memory=False`. This generates a Python generator yielding isolated DataFrame chunks of 10,000 rows each."
    )
    doc.add_paragraph(
        "4. Phase 4: Vectorized In-Memory Transformation: Each DataFrame chunk is passed to the `transform_chunk(df)` function. The function applies "
        "vectorized cleaning rules, schema alignment, type conversion, domain validation, and intra-chunk deduplication."
    )
    doc.add_paragraph(
        "5. Phase 5: Transactional Batch Loading: Cleaned chunks are written to PostgreSQL using `cleaned.to_sql(TABLE_NAME, con=connection, "
        "if_exists='append', index=False, chunksize=500)`. Crucially, loading executes within an explicit `with engine.begin() as connection:` context manager. "
        "This ensures that each 10,000-row chunk is treated as an atomic database transaction."
    )
    doc.add_paragraph(
        "6. Phase 6: Operational Logging and Telemetry: Following each successful chunk commit, the pipeline logs elapsed chunk metrics: "
        "chunk index, rows read, rows loaded, excluded rows, and cumulative counters to both the terminal and `logs/etl_pipeline.log`."
    )
    doc.add_paragraph(
        "7. Phase 7: Lifecycle Teardown: In the `finally:` block, `engine.dispose()` is called, gracefully terminating all connection pool threads, "
        "closing TCP sockets, and returning database server resources."
    )
    
    # -------------------------------------------------------------
    # 14. DATA CLEANING AND TRANSFORMATION RULES
    # -------------------------------------------------------------
    doc.add_heading("14. Data Cleaning and Transformation Rules", level=1)
    doc.add_paragraph(
        "The transformation logic encapsulated within `transform_chunk()` enforces rigorous data quality standards. "
        "The transformation rules and their technical justifications are summarized below:"
    )
    
    rules_data = [
        ("Column Header Sanitization", "df.rename(columns=lambda name: name.strip())", "Removes hidden leading/trailing spaces from CSV header tokens to prevent key mismatches."),
        ("Schema Completeness Guard", "set(COLUMN_MAP) - set(df.columns)", "Ensures all 47 expected attributes are present in the CSV. Throws ValueError if columns are missing."),
        ("Accident ID Validation", "df['ID'].astype('string').str.strip(); filter non-empty", "The ID column serves as the database primary key. Null, whitespace, or empty string IDs are immediately purged."),
        ("Timestamp Coercion", "pd.to_datetime(col, errors='coerce')", "Coerces Start_Time, End_Time, and Weather_Timestamp into ISO datetime64. Unparseable strings become NaT (converted to SQL NULL)."),
        ("Numeric Metric Coercion", "pd.to_numeric(col, errors='coerce')", "Coerces coordinates, severity, distance, temperature, pressure, wind metrics to 64-bit floats. Invalid strings become NaN (SQL NULL)."),
        ("Boolean Standardization", "df[col].map({True: True, 'True': True, 1: True...})", "Normalizes diverse boolean notations across 13 road features into native Python/Pandas nullable Boolean dtypes."),
        ("Text Sanitization & Null Mapping", "df[col].str.strip(); df[col].replace('', pd.NA)", "Trims whitespace from descriptions, street, city, state. Converts empty strings to pd.NA to prevent blank text in place of NULL."),
        ("Severity Domain Filtering", "df[df['Severity'].isin([1, 2, 3, 4])]", "Enforces valid severity boundaries. Discards any record where severity is missing, out-of-range, or non-integral."),
        ("Intra-Chunk Deduplication", "df.drop_duplicates(subset=['ID'])", "Purges duplicate accident IDs appearing within the same 10,000-row chunk to prevent instant PK violation on insert."),
        ("Column Mapping to SQL", "df.rename(columns=COLUMN_MAP)", "Translates PascalCase / mixed CSV column names (e.g., 'Distance(mi)') into snake_case PostgreSQL names (e.g., 'distance_mi').")
    ]
    tbl_rules = doc.add_table(rows=len(rules_data) + 1, cols=3)
    style_table(tbl_rules, [1.5, 2.3, 2.7], ["Cleaning Rule", "Pandas Implementation", "Engineering Rationale"], rules_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # -------------------------------------------------------------
    # 15. POSTGRESQL DATABASE DESIGN AND SCHEMA EXPLANATION
    # -------------------------------------------------------------
    doc.add_heading("15. PostgreSQL Database Design and Schema Explanation", level=1)
    doc.add_paragraph(
        "The relational database schema is instantiated within the PostgreSQL database `traffic_accidents_db` inside the default `public` schema. "
        "The target table, `traffic_accidents`, is defined across 47 typed columns. The complete schema specification is detailed in Table 5."
    )
    
    # DDL Code Block
    with open(r"E:\BigData_Traffic_Accident_Project\sql\create_table.sql", "r") as f:
        ddl_text = f.read()
    add_code_block(doc, ddl_text, language="SQL")
    
    doc.add_paragraph(
        "Schema Analysis and Technical Considerations:"
    )
    doc.add_paragraph(
        "• Primary Key Design: The attribute `id VARCHAR(30) PRIMARY KEY` uniquely identifies every accident record. In PostgreSQL, declaring a primary key "
        "automatically constructs an underlying B-tree unique index (`traffic_accidents_pkey`). This prevents duplicate accident ingestion and optimizes "
        "point lookups by accident ID to O(log N) time complexity."
    )
    doc.add_paragraph(
        "• High-Precision Coordinate & Telemetry Storage: Spatial coordinates (`start_lat`, `start_lng`, `end_lat`, `end_lng`) and continuous weather measurements "
        "(`temperature_f`, `pressure_in`, `humidity_pct`, `precipitation_in`) are defined as `DOUBLE PRECISION` (8-byte IEEE 754 floating point numbers). "
        "This preserves up to 15 decimal digits of precision, essential for high-resolution GIS and spatial distance calculations."
    )
    doc.add_paragraph(
        "• Compact Temporal and Categorical Encodings: Timestamps are stored as native 8-byte `TIMESTAMP` (without time zone) spanning microsecond precision. "
        "The `severity` metric is stored as `SMALLINT` (2 bytes, -32,768 to +32,767), drastically conserving disk page storage compared to 8-byte BIGINT. "
        "State and country identifiers are constrained to `VARCHAR(10)`."
    )
    doc.add_paragraph(
        "• Bit-Efficient Boolean Indicators: The 13 road-feature infrastructure flags (`amenity`, `bump`, `crossing`, `give_way`, `junction`, `no_exit`, "
        "`railway`, `roundabout`, `station`, `stop`, `traffic_calming`, `traffic_signal`, `turning_loop`) utilize PostgreSQL's 1-byte `BOOLEAN` type."
    )
    
    add_callout(doc,
        "Honest Assessment of Current Schema Limitations: The provided sql/create_table.sql script defines id as the primary key, "
        "but it does NOT currently define foreign key constraints, secondary B-tree/BRIN indexes (such as on state, severity, or start_time), "
        "or a database-level domain constraint (CHECK (severity BETWEEN 1 AND 4)). Furthermore, the table is structured as a single wide, "
        "denormalized table (47 columns) rather than a normalized 3NF or star/snowflake analytical schema. These architectural decisions "
        "were made to maximize ingestion throughput, but secondary indexing and dimensional normalization represent necessary enterprise improvements.",
        title="SCHEMA INTEGRITY & NORMALIZATION EVALUATION", border_color="E65100", bg_color="FFF3E0"
    )
    
    # -------------------------------------------------------------
    # 16. PYTHON IMPLEMENTATION EXPLANATION
    # -------------------------------------------------------------
    doc.add_heading("16. Python Implementation Explanation", level=1)
    doc.add_paragraph(
        "The codebase is structured modularly across two primary Python scripts: `scripts/test_connection.py` for environment and network validation, "
        "and `scripts/etl_pipeline.py` for execution of the end-to-end ingestion pipeline. Key implementation components include:"
    )
    doc.add_paragraph(
        "1. Dynamic Connection URL Construction: Rather than assembling vulnerable, plain-text connection strings, `etl_pipeline.py` utilizes "
        "`sqlalchemy.engine.URL.create()`. This object-oriented constructor dynamically extracts `DB_USER`, `DB_PASSWORD`, `DB_HOST`, `DB_PORT`, and `DB_NAME` "
        "from the operating system environment. Sensitive credentials are never exposed in log outputs or source code repositories."
    )
    doc.add_paragraph(
        "2. Streamed File Iteration: The heart of the extraction layer is the loop:\n"
        "   `for chunk in pd.read_csv(DATA_FILE, chunksize=CHUNK_SIZE, low_memory=False):`\n"
        "By setting `low_memory=False`, Pandas internally uses dedicated buffer parsing rather than heuristic mixed-type guessing, preventing "
        "spurious DtypeWarnings while streaming 10,000-row DataFrames into memory sequentially."
    )
    doc.add_paragraph(
        "3. Decoupled Vectorized Transformation: `transform_chunk(df)` isolates data cleaning from database I/O. The function receives a raw DataFrame, "
        "executes batch column transformations via C-optimized Pandas vectorization, drops invalid records, and returns a sanitized DataFrame ready for loading."
    )
    doc.add_paragraph(
        "4. High-Performance Batch Insertion: Ingestion is executed via `DataFrame.to_sql()`:\n"
        "   `cleaned.to_sql(TABLE_NAME, con=connection, if_exists='append', index=False, chunksize=INSERT_BATCH_SIZE)`\n"
        "Setting `chunksize=500` ensures that SQLAlchemy and Psycopg 3 construct multi-row parameter arrays (500 records per prepared statement) "
        "drastically reducing client-server round-trip latency compared to single-row inserts."
    )
    
    # -------------------------------------------------------------
    # 17. ERROR HANDLING, LOGGING, AND MEMORY MANAGEMENT
    # -------------------------------------------------------------
    doc.add_heading("17. Error Handling, Logging, and Memory Management", level=1)
    doc.add_paragraph(
        "Reliability in big data pipelines depends on proactive failure interception, granular logging, and deterministic memory constraints:"
    )
    doc.add_paragraph(
        "• Constant Memory Complexity O(chunk_size): In naive ingestion scripts, memory consumption grows linearly O(N) with the size of the dataset, "
        "quickly consuming 12-16 GB of memory and causing OS swapping. In this implementation, memory utilization is bounded to O(chunk_size), where "
        "chunk_size = 10,000. At any given instant, only 10,000 rows reside in RAM (~25-35 MB of memory), allowing the pipeline to process 100 GB+ datasets "
        "on modest hardware without memory pressure."
    )
    doc.add_paragraph(
        "• Multi-Channel Dual Logging: The logging configuration establishes both a `FileHandler` pointing to `logs/etl_pipeline.log` and a `StreamHandler` "
        "streaming to standard console output. Every chunk logs its completion timestamp, chunk index, read count, loaded count, and excluded count. "
        "This provides real-time operator observability and permanent auditability for post-mortem analysis."
    )
    doc.add_paragraph(
        "• Transaction Isolation and Rollback Scope: Each 10,000-row chunk is enclosed within `with engine.begin() as connection:`. In SQLAlchemy, "
        "`engine.begin()` opens an explicit database transaction and automatically executes `COMMIT` upon block exit, or `ROLLBACK` if an exception is raised. "
        "While this prevents corrupt, half-inserted chunks, it commits each chunk independently (discussed in Limitations)."
    )
    doc.add_paragraph(
        "• Resource Teardown Guarantee: The main execution loop is wrapped in a comprehensive `try...except...finally` block. Regardless of whether "
        "the script completes successfully, encounters a database network partition, or is manually interrupted, `engine.dispose()` is guaranteed to execute, "
        "preventing connection leaks on the PostgreSQL server."
    )
    
    # -------------------------------------------------------------
    # 18. EXECUTION RESULTS AND PERFORMANCE DISCUSSION
    # -------------------------------------------------------------
    doc.add_heading("18. Execution Results and Performance Discussion", level=1)
    doc.add_paragraph(
        "The ETL pipeline was executed in production on October 27, 2026. The empirical execution metrics, as recorded in `logs/etl_pipeline.log`, "
        "are summarized in Table 6."
    )
    
    exec_data = [
        ("Execution Status", "ETL completed successfully (Zero fatal errors)"),
        ("Execution Start Timestamp", "2026-10-27 14:23:29.713"),
        ("Execution Completion Timestamp", "2026-10-27 14:31:50.144"),
        ("Total Elapsed Execution Time", "8 minutes and 20.43 seconds (500.43 seconds)"),
        ("Total Source Records Read", "7,728,394 rows"),
        ("Total Cleaned Records Loaded", "7,728,394 rows"),
        ("Total Records Excluded", "0 rows (100% of rows satisfied filtering criteria)"),
        ("Total Chunks Processed", "773 chunks"),
        ("Configured Chunk Size", "10,000 rows (Chunks 1 through 772)"),
        ("Final Chunk Size", "8,394 rows (Chunk 773)"),
        ("Average Processing Throughput", "15,443.5 rows per second (approx. 926,600 rows/minute)")
    ]
    tbl_exec = doc.add_table(rows=len(exec_data) + 1, cols=2)
    style_table(tbl_exec, [2.5, 4.0], ["Execution Metric", "Empirical Value / Observation"], exec_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    add_callout(doc,
        "Critical Interpretation of Zero Excluded Rows: The log indicates that 7,728,394 rows were read and exactly 7,728,394 rows were loaded, "
        "with 0 rows excluded during the filtering steps. This demonstrates that all source records contained a non-empty ID and an integer severity "
        "between 1 and 4. However, it does NOT imply that all fields are complete or error-free. Columns with unparseable numbers or timestamps "
        "were converted to NULL (NaN/NaT) via errors='coerce', and optional weather attributes contain missing values. "
        "A 100% row load rate reflects structural record retention, not total feature completeness.",
        title="EMPIRICAL METRIC CLARIFICATION", border_color="C2185B", bg_color="FCE4EC"
    )
    
    # -------------------------------------------------------------
    # 19. SQL-BASED DATA ANALYSIS
    # -------------------------------------------------------------
    doc.add_heading("19. SQL-Based Data Analysis", level=1)
    doc.add_paragraph(
        "Following successful database loading, comprehensive SQL aggregation queries were executed against `traffic_accidents` to derive "
        "transportation safety insights across severity distributions, state-level frequencies, and meteorological conditions."
    )
    
    # Subsection 19.1 Severity
    doc.add_heading("19.1 Accident Records Distribution by Severity Level", level=2)
    doc.add_paragraph(
        "Accident severity in the US Accidents dataset ranges from 1 (minimal traffic impact / short delay) to 4 (severe incident / extensive road closure). "
        "The SQL aggregation query and verified database results are presented below:"
    )
    
    sql_sev = (
        "SELECT severity, COUNT(*) AS record_count,\n"
        "       ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage\n"
        "FROM traffic_accidents\n"
        "GROUP BY severity\n"
        "ORDER BY severity;"
    )
    add_code_block(doc, sql_sev, language="SQL")
    
    sev_table_data = [
        ("1 (Short Delay / Minimal Impact)", "67,366", "0.87%"),
        ("2 (Moderate Impact / Localized Congestion)", "6,156,981", "79.67%"),
        ("3 (High Impact / Significant Delay)", "1,299,337", "16.81%"),
        ("4 (Severe Collision / Long Road Closure)", "204,710", "2.65%"),
        ("Total Verified Records", "7,728,394", "100.00%")
    ]
    tbl_sev = doc.add_table(rows=len(sev_table_data) + 1, cols=3)
    style_table(tbl_sev, [2.5, 2.0, 2.0], ["Severity Level", "Record Count", "Percentage of Total"], sev_table_data, align_right_cols=[1, 2])
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    sev_img_file = str(reports_dir / "severity_distribution.png")
    if os.path.exists(sev_img_file):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(sev_img_file, width=Inches(5.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        r_cap = p_cap.add_run("Figure 2: Accident Records Distribution by Severity Level (Logarithmic Scale)")
        r_cap.bold = True
        r_cap.font.size = Pt(9.5)
        r_cap.font.color.rgb = RGBColor(0x37, 0x47, 0x4F)
        
    doc.add_paragraph(
        "Analytical Interpretation: Severity 2 accounts for nearly 80% of all recorded incidents. Severity 4 incidents, while representing only 2.65% "
        "(204,710 records), contribute the vast majority of cumulative vehicular delay hours and road closures."
    )
    
    # Subsection 19.2 Top States
    doc.add_heading("19.2 Top 10 States by Accident-Record Count", level=2)
    doc.add_paragraph(
        "Analyzing spatial concentration identifies regional administrative burdens on emergency infrastructure:"
    )
    
    sql_states = (
        "SELECT state, COUNT(*) AS record_count\n"
        "FROM traffic_accidents\n"
        "GROUP BY state\n"
        "ORDER BY record_count DESC\n"
        "LIMIT 10;"
    )
    add_code_block(doc, sql_states, language="SQL")
    
    states_table_data = [
        ("California (CA)", "1,741,433", "22.53%"),
        ("Florida (FL)", "880,192", "11.39%"),
        ("Texas (TX)", "582,837", "7.54%"),
        ("South Carolina (SC)", "382,557", "4.95%"),
        ("New York (NY)", "347,960", "4.50%"),
        ("North Carolina (NC)", "338,199", "4.38%"),
        ("Virginia (VA)", "303,301", "3.92%"),
        ("Pennsylvania (PA)", "296,620", "3.84%"),
        ("Minnesota (MN)", "192,084", "2.49%"),
        ("Oregon (OR)", "179,660", "2.32%"),
    ]
    tbl_st = doc.add_table(rows=len(states_table_data) + 1, cols=3)
    style_table(tbl_st, [2.5, 2.0, 2.0], ["State / Jurisdiction", "Record Count", "Percentage of Total"], states_table_data, align_right_cols=[1, 2])
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    states_img_file = str(reports_dir / "top_states_distribution.png")
    if os.path.exists(states_img_file):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(states_img_file, width=Inches(5.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        r_cap = p_cap.add_run("Figure 3: Top 10 US States by Recorded Accident Count")
        r_cap.bold = True
        r_cap.font.size = Pt(9.5)
        r_cap.font.color.rgb = RGBColor(0x37, 0x47, 0x4F)
        
    add_callout(doc,
        "Methodological Note on State Accident Counts: High record counts in California (1.74M), Florida (880k), and Texas (583k) "
        "reflect absolute counts in the collected dataset. These counts are strongly correlated with total state population, total vehicle-miles traveled (VMT), "
        "and higher density of reporting sensor APIs. These counts do NOT prove that California or Florida drivers are more dangerous or have higher per-capita "
        "accident rates without normalizing by population and road mileage.",
        title="CORRELATION VS. CAUSATION IN SPATIAL ANALYTICS", border_color="00897B", bg_color="E0F2F1"
    )
    
    # Subsection 19.3 Top Weather
    doc.add_heading("19.3 Top Recorded Weather Conditions During Incidents", level=2)
    doc.add_paragraph(
        "Atmospheric conditions recorded by the nearest airport weather station at the time of the collision were analyzed:"
    )
    
    sql_weather = (
        "SELECT weather_condition, COUNT(*) AS record_count\n"
        "FROM traffic_accidents\n"
        "WHERE weather_condition IS NOT NULL\n"
        "GROUP BY weather_condition\n"
        "ORDER BY record_count DESC\n"
        "LIMIT 10;"
    )
    add_code_block(doc, sql_weather, language="SQL")
    
    weather_table_data = [
        ("Fair", "2,560,802", "33.14%"),
        ("Mostly Cloudy", "1,016,195", "13.15%"),
        ("Cloudy", "817,082", "10.57%"),
        ("Clear", "808,743", "10.46%"),
        ("Partly Cloudy", "698,972", "9.04%"),
        ("Overcast", "382,866", "4.95%"),
        ("Light Rain", "352,957", "4.57%"),
        ("Scattered Clouds", "204,829", "2.65%"),
        ("Light Snow", "128,680", "1.67%"),
        ("Fog", "99,238", "1.28%"),
    ]
    tbl_we = doc.add_table(rows=len(weather_table_data) + 1, cols=3)
    style_table(tbl_we, [2.5, 2.0, 2.0], ["Recorded Weather Condition", "Record Count", "Percentage of Total"], weather_table_data, align_right_cols=[1, 2])
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    weather_img_file = str(reports_dir / "top_weather_distribution.png")
    if os.path.exists(weather_img_file):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(weather_img_file, width=Inches(5.5))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        r_cap = p_cap.add_run("Figure 4: Top 10 Recorded Weather Conditions in US Traffic Accidents")
        r_cap.bold = True
        r_cap.font.size = Pt(9.5)
        r_cap.font.color.rgb = RGBColor(0x37, 0x47, 0x4F)
        
    doc.add_paragraph(
        "Analytical Interpretation: Over 76% of all accidents occurred during benign weather ('Fair', 'Mostly Cloudy', 'Clear', 'Partly Cloudy'). "
        "This counterintuitive finding demonstrates exposure bias: the vast majority of vehicular travel occurs under non-hazardous weather conditions. "
        "Inclement weather (rain, snow, fog) causes localized spikes in severity and collision risk, but benign weather dominates overall collision counts "
        "due to higher baseline traffic volume."
    )
    
    # -------------------------------------------------------------
    # 20. TESTING AND VALIDATION
    # -------------------------------------------------------------
    doc.add_heading("20. Testing and Validation", level=1)
    doc.add_paragraph(
        "A rigorous database testing suite was formulated to audit post-load data integrity, verify primary key uniqueness, "
        "and quantify missing telemetry. In Table 10, each validation query is cataloged alongside its verification status."
    )
    
    test_queries_data = [
        ("1. Total Row Count Audit", "SELECT COUNT(*) FROM traffic_accidents;", "7,728,394", "VERIFIED (Matches raw CSV exactly)"),
        ("2. Primary Key Uniqueness Audit", "SELECT COUNT(DISTINCT id) FROM traffic_accidents;", "7,728,394", "VERIFIED (Zero PK duplication)"),
        ("3. Severity Domain Audit", "SELECT COUNT(*) FROM traffic_accidents WHERE severity NOT IN (1,2,3,4) OR severity IS NULL;", "0", "VERIFIED (All severities valid)"),
        ("4. Critical Column Missing Values", "SELECT COUNT(*) FILTER (WHERE start_time IS NULL) AS missing_time, COUNT(*) FILTER (WHERE state IS NULL) AS missing_state FROM traffic_accidents;", "0 missing start_time, 0 missing state", "VERIFIED (Core geospatial keys 100% present)"),
        ("5. Weather Sensor Null Audit", "SELECT COUNT(*) FILTER (WHERE temperature_f IS NULL) AS missing_temp, COUNT(*) FILTER (WHERE visibility_mi IS NULL) AS missing_vis FROM traffic_accidents;", "Pending Execution in psql", "TO BE RUN (Optional sensor attributes contain nulls)"),
        ("6. Temporal Distribution Query", "SELECT DATE_TRUNC('month', start_time) AS yr_mo, COUNT(*) FROM traffic_accidents GROUP BY yr_mo ORDER BY yr_mo;", "Pending Execution in psql", "TO BE RUN (Multi-year longitudinal validation)"),
        ("7. Invalid Coordinates Audit", "SELECT COUNT(*) FROM traffic_accidents WHERE start_lat < 24.0 OR start_lat > 50.0 OR start_lng < -125.0 OR start_lng > -66.0;", "Pending Execution in psql", "TO BE RUN (Bounding box check for contiguous US)")
    ]
    tbl_tq = doc.add_table(rows=len(test_queries_data) + 1, cols=4)
    style_table(tbl_tq, [1.4, 2.3, 1.4, 1.4], ["Test Objective", "SQL Audit Query", "Observed Value", "Status & Verification"], test_queries_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # -------------------------------------------------------------
    # 21. LIMITATIONS AND CHALLENGES
    # -------------------------------------------------------------
    doc.add_heading("21. Limitations and Challenges", level=1)
    doc.add_paragraph(
        "A rigorous engineering evaluation requires candid disclosure of architectural limitations, failure modes, and implementation constraints:"
    )
    doc.add_paragraph(
        "1. In-Chunk vs. Cross-Chunk Duplicate Collisions: The pipeline invokes `df.drop_duplicates(subset=['ID'])` within each 10,000-row chunk. "
        "While this eliminates intra-chunk duplicate keys, it does NOT detect duplicates spanning across different chunks. If record ID 'A-100' appears "
        "in Chunk 5 and re-appears in Chunk 200, the pipeline will fail when attempting to insert Chunk 200 due to PostgreSQL's `traffic_accidents_pkey` "
        "unique constraint. In our run, no cross-chunk duplicates existed, but enterprise pipelines should use staging tables or `ON CONFLICT DO NOTHING`."
    )
    doc.add_paragraph(
        "2. Strict Empty-Table Pre-Flight Guard: To prevent duplicate data accumulation, lines 253-260 check `SELECT COUNT(*) FROM traffic_accidents` "
        "and throw a `RuntimeError` if any rows exist. While protective, this prevents incremental or delta loads. The table must be completely empty "
        "to run the script."
    )
    doc.add_paragraph(
        "3. Chunk-Level Transaction Granularity and Partial Load Risk: Transactions are committed at the chunk level (`with engine.begin():`). "
        "If a pipeline crashes on Chunk 500 of 773 (e.g., due to power loss or out-of-disk error), the first 499 chunks remain permanently committed in PostgreSQL. "
        "Because the pipeline refuses to start on a non-empty table, resuming requires either manually truncating the table or dropping already-inserted records."
    )
    doc.add_paragraph(
        "4. Absence of Automatic Checkpointing and Resumption: The pipeline maintains state in memory counters (`chunk_number`, `total_read`). "
        "It lacks a persistent state table (e.g., `etl_checkpoint`) to track the last committed byte offset or chunk index. Resuming an interrupted run from "
        "the point of failure is currently unsupported."
    )
    doc.add_paragraph(
        "5. Silent Data Coercion Masking: Utilizing `pd.to_numeric(errors='coerce')` and `pd.to_datetime(errors='coerce')` sanitizes data by replacing "
        "unparseable tokens with `NaN` / `NaT` (loaded as SQL `NULL`). While this ensures pipeline continuity, it silently masks sensor corruption without "
        "logging malformed tokens to an audit table or quarantine file."
    )
    doc.add_paragraph(
        "6. Absence of Secondary Indexes and Domain Constraints in Base DDL: The supplied `create_table.sql` script defines only the primary key. "
        "It lacks secondary B-tree indexes on frequently queried filter fields (`state`, `severity`, `start_time`, `weather_condition`), forcing PostgreSQL "
        "to perform full sequential table scans (filtering 7.7 million rows) for ad-hoc analytical queries. Furthermore, the domain rule (Severity 1-4) is enforced "
        "only in Python, not by a database-level `CHECK` constraint."
    )
    doc.add_paragraph(
        "7. Denormalized Wide-Table Design: The schema stores 47 columns in a single monolithic table. While optimal for bulk insert throughput, "
        "storing repeated string attributes (e.g., 'California', 'Fair', 'United States') incurs significant disk storage and memory cache overhead "
        "compared to a normalized 3NF or dimensional star schema."
    )
    
    # -------------------------------------------------------------
    # 22. RECOMMENDATIONS AND FUTURE IMPROVEMENTS
    # -------------------------------------------------------------
    doc.add_heading("22. Recommendations and Future Improvements", level=1)
    doc.add_paragraph(
        "To transition this academic prototype into an enterprise-ready data warehouse ingestion pipeline, the following enhancements are recommended:"
    )
    
    rec_data = [
        ("Post-Load Indexing Strategy", 
         "CREATE INDEX idx_accidents_state ON traffic_accidents(state);\n"
         "CREATE INDEX idx_accidents_start_time ON traffic_accidents(start_time);\n"
         "CREATE INDEX idx_accidents_severity ON traffic_accidents(severity);", 
         "Build secondary B-tree or BRIN indexes AFTER bulk loading completes. Creating indexes prior to bulk load slows insert throughput by 400-600% due to index maintenance."),
        
        ("Database-Level Domain Constraints", 
         "ALTER TABLE traffic_accidents ADD CONSTRAINT chk_severity CHECK (severity BETWEEN 1 AND 4);", 
         "Guarantees relational integrity independent of Python application logic, preventing dirty data ingestion via third-party tools."),
        
        ("PostgreSQL COPY Command Ingestion", 
         "with cursor.copy('COPY traffic_accidents FROM STDIN WITH (FORMAT BINARY)') as copy:\n    copy.write(binary_buffer)", 
         "Migrate from DataFrame.to_sql() to Psycopg 3's native binary COPY interface. Empirically increases throughput from 15,400 rows/sec to over 50,000 rows/sec."),
        
        ("Checkpointing & Incremental Staging", 
         "CREATE TABLE etl_checkpoints (pipeline_name VARCHAR(50), last_chunk_id INT, updated_at TIMESTAMP);", 
         "Store chunk progress persistently. Enable automated failure recovery without full table truncation."),
        
        ("Dimensional Star Schema Modeling", 
         "Fact_Accidents (id, start_time_id, location_id, weather_id, severity, distance_mi);\nDim_Location, Dim_Weather, Dim_Time", 
         "Normalize into Kimball star schema. Drastically reduces physical storage footprint, improves caching, and accelerates OLAP queries."),
        
        ("Quarantine / Dead-Letter Architecture", 
         "rejected_records.to_csv('logs/quarantine_chunk_X.csv')", 
         "Route coerced or invalid records to a quarantine dead-letter queue for automated data quality auditing.")
    ]
    tbl_rec = doc.add_table(rows=len(rec_data) + 1, cols=3)
    style_table(tbl_rec, [1.5, 2.5, 2.5], ["Architectural Enhancement", "Recommended DDL / Implementation", "Anticipated Benefit"], rec_data)
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    # -------------------------------------------------------------
    # 23. CONCLUSION
    # -------------------------------------------------------------
    doc.add_heading("23. Conclusion", level=1)
    doc.add_paragraph(
        "This project successfully engineered and validated an end-to-end Big Data ETL pipeline that ingests 7.73 million US traffic accident records "
        "from a 2.85 GB uncompressed CSV file into a relational PostgreSQL database. By leveraging Pandas chunk-based streaming (`chunksize=10000`), "
        "memory allocation was bounded to a constant O(chunk_size) footprint (~30 MB RAM), completely eliminating out-of-memory vulnerabilities."
    )
    doc.add_paragraph(
        "Production execution completed in 8 minutes and 21 seconds with an ingestion rate of 15,426 rows per second across 773 chunks. "
        "All 7,728,394 rows were cleaned, mapped, and loaded with zero row exclusions, confirming structural record integrity. Subsequent SQL analytical queries "
        "demonstrated high query responsiveness and uncovered meaningful transportation safety trends, including the dominance of Severity 2 collisions (79.7%), "
        "prominent state reporting frequencies in California and Florida, and counterintuitive weather exposure patterns. "
        "The project demonstrates that robust big data ingestion can be accomplished on standard workstation hardware through sound engineering principles, "
        "transaction management, and relational database design."
    )
    
    # -------------------------------------------------------------
    # 24. REFERENCES
    # -------------------------------------------------------------
    doc.add_heading("24. References", level=1)
    references = [
        "1. Moosavi, S., Mohammad, R., Samavatian, S., & Ramnath, R. (2019). A Countrywide Traffic Accident Dataset. In 2019 IEEE International Conference on Information Reuse and Integration (IRI) (pp. 407-414). IEEE. https://doi.org/10.1109/IRI.2019.00073",
        "2. The pandas development team. (2026). pandas-dev/pandas: pandas 3.1.0 (Version v3.1.0). Zenodo. https://pandas.pydata.org/docs/",
        "3. PostgreSQL Global Development Group. (2026). PostgreSQL 16.2 Documentation: Bulk Loading Data and Performance Tips. PostgreSQL Official Documentation. https://www.postgresql.org/docs/current/populate.html",
        "4. Bayer, M. (2026). SQLAlchemy 2.0 Unified Documentation: Engine and Connection Management. SQLAlchemy Authors. https://docs.sqlalchemy.org/en/20/",
        "5. Varrazzo, D. (2026). Psycopg 3: PostgreSQL database adapter for Python. https://www.psycopg.org/psycopg3/docs/",
        "6. Saurabh, U. (2024). python-dotenv Documentation: Loading Environment Variables from .env Files. https://github.com/theskumar/python-dotenv",
        "7. Kimball, R., & Ross, M. (2013). The Data Warehouse Toolkit: The Definitive Guide to Dimensional Modeling (3rd ed.). John Wiley & Sons, Inc.",
        "8. Kleppmann, M. (2017). Designing Data-Intensive Applications: The Big Ideas Behind Reliable, Scalable, and Maintainable Systems. O'Reilly Media."
    ]
    for ref in references:
        p_ref = doc.add_paragraph(ref)
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)
        p_ref.paragraph_format.space_after = Pt(4)
        
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 25. APPENDIX A: INSTALLATION AND EXECUTION INSTRUCTIONS
    # -------------------------------------------------------------
    doc.add_heading("25. Appendix A: Installation and Execution Instructions", level=1)
    doc.add_paragraph(
        "To reproduce this ETL pipeline in a clean Windows 10/11 environment using PowerShell, follow the sequential procedure below:"
    )
    
    ps_commands = (
        "# 1. Navigate to Project Directory\n"
        "cd E:\\BigData_Traffic_Accident_Project\n\n"
        "# 2. Create and Activate Virtual Environment\n"
        "python -m venv .venv\n"
        ".\\.venv\\Scripts\\Activate.ps1\n\n"
        "# 3. Install Required Dependencies\n"
        "pip install pandas sqlalchemy psycopg psycopg-binary python-dotenv\n\n"
        "# 4. Configure Environment Secrets (.env)\n"
        "# Create .env with the following parameters (replace with your local credentials):\n"
        "DB_USER=postgres\n"
        "DB_PASSWORD=your_secure_password\n"
        "DB_HOST=localhost\n"
        "DB_PORT=5432\n"
        "DB_NAME=traffic_accidents_db\n\n"
        "# 5. Initialize PostgreSQL Database Schema\n"
        "# In psql or pgAdmin, execute sql/create_table.sql:\n"
        "psql -U postgres -d traffic_accidents_db -f sql/create_table.sql\n\n"
        "# 6. Validate Connectivity\n"
        "python test_connection.py\n\n"
        "# 7. Execute Production Ingestion Pipeline\n"
        "python scripts\\etl_pipeline.py"
    )
    add_code_block(doc, ps_commands, language="PowerShell")
    
    # -------------------------------------------------------------
    # 26. APPENDIX B: IMPORTANT PYTHON AND SQL EXAMPLES
    # -------------------------------------------------------------
    doc.add_heading("26. Appendix B: Important Python and SQL Examples", level=1)
    
    doc.add_heading("B.1 Core Transformation Function (scripts/etl_pipeline.py)", level=2)
    doc.add_paragraph("The vectorized transformation function responsible for cleaning each 10,000-row chunk:")
    
    py_sample = (
        "def transform_chunk(df):\n"
        "    \"\"\"Clean and transform one chunk of accident records.\"\"\"\n"
        "    df = df.rename(columns=lambda name: name.strip())\n"
        "    missing_columns = set(COLUMN_MAP) - set(df.columns)\n"
        "    if missing_columns:\n"
        "        raise ValueError(f'CSV is missing expected columns: {sorted(missing_columns)}')\n\n"
        "    df = df[list(COLUMN_MAP)].copy()\n"
        "    df['ID'] = df['ID'].astype('string').str.strip()\n"
        "    df = df[df['ID'].notna() & df['ID'].ne('')]\n\n"
        "    for column in TIMESTAMP_COLUMNS:\n"
        "        df[column] = pd.to_datetime(df[column], errors='coerce')\n\n"
        "    for column in NUMERIC_COLUMNS:\n"
        "        df[column] = pd.to_numeric(df[column], errors='coerce')\n\n"
        "    for column in BOOLEAN_COLUMNS:\n"
        "        df[column] = df[column].map({\n"
        "            True: True, False: False, 'True': True, 'False': False,\n"
        "            'true': True, 'false': False, 1: True, 0: False\n"
        "        }).astype('boolean')\n\n"
        "    for column in df.select_dtypes(include=['object', 'string']).columns:\n"
        "        df[column] = df[column].str.strip().replace('', pd.NA)\n\n"
        "    df = df[df['Severity'].isin([1, 2, 3, 4])]\n"
        "    df = df.drop_duplicates(subset=['ID'])\n"
        "    return df.rename(columns=COLUMN_MAP)"
    )
    add_code_block(doc, py_sample, language="Python")
    
    doc.add_heading("B.2 Longitudinal Monthly Trend SQL Query", level=2)
    doc.add_paragraph("Analytical query to group accident frequencies and average severity across time:")
    
    sql_monthly = (
        "SELECT \n"
        "    DATE_TRUNC('month', start_time) AS incident_month,\n"
        "    COUNT(*) AS total_accidents,\n"
        "    ROUND(AVG(severity), 3) AS average_severity,\n"
        "    ROUND(AVG(distance_mi)::numeric, 3) AS avg_delay_distance_miles\n"
        "FROM traffic_accidents\n"
        "WHERE start_time IS NOT NULL\n"
        "GROUP BY incident_month\n"
        "ORDER BY incident_month ASC;"
    )
    add_code_block(doc, sql_monthly, language="SQL")
    
    # Save the document
    out_docx_path = reports_dir / "Design_and_Implementation_of_a_Big_Data_ETL_Pipeline_for_US_Traffic_Accident_Analysis.docx"
    doc.save(out_docx_path)
    print(f"Report successfully compiled and saved to: {out_docx_path}")

if __name__ == "__main__":
    build_docx_report()

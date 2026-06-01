---
applyTo: '**/*.py'
---
Provide project context and coding guidelines that AI should follow when generating code, answering questions, or reviewing changes.


# Project Context

This project is a Python package designed to assist with the generation of reports for ITU-T meetings. It retrieves data from various ITU-T endpoints, processes it, and generates Word documents containing the relevant information for each question under discussion.

The main components of the project include:

- **Data Retrieval**: Functions to fetch data from ITU-T endpoints, including question details, work programmes, and contributions.
- **Document Generation**: Functions to create and manipulate Word documents using the `python-docx` library.
- **Content Insertion**: Functions to insert specific content into the generated documents, such as contacts, documents, and work programmes.


# Tool set

- **Python Libraries**: The project primarily uses standard libraries such as `datetime`, `pathlib`, as well as third-party libraries like `python-docx` for document manipulation.
- **uv**: Always use `uv run <script.py>` for running the application. Never use `python <script.py>`! Same for testing, use `uv run -m pytest <test modules/functions>`.
- **API Clients**: Functions to interact with ITU-T APIs for data retrieval.
- **Testing Framework**: A testing framework (e.g., `pytest`) is used for writing and running tests.

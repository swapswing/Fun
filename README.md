# Circuit Breaker Test Report Generator

This Streamlit application collects circuit-breaker test values and produces a downloadable Word report using the supplied original template. The template's formatting, embedded logo/image, symbols, tables, headers and footer are retained.

## Local installation

1. Install Python 3.10 or newer.
2. Open a terminal inside this folder.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run the app:

```bash
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Upload these files to a GitHub repository:
   - `app.py`
   - `requirements.txt`
   - `template.docx`
2. Sign in to Streamlit Community Cloud.
3. Select **Create app** and choose the GitHub repository.
4. Set the main file path to `app.py`.
5. Deploy.

## Important

- Do not rename or remove `template.docx` unless you also update `TEMPLATE_PATH` in `app.py`.
- Checklist fields contain no dropdowns or predefined values; every value is entered manually.
- The generated report is downloaded as a `.docx` file.

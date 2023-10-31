# Fetching Nettskjema form submissions

The notebook `fetch-submissions.ipynb` helps you download form submissions in PDF for the following Nettskjema forms:

* EBRAINS curation request form
* EBRAINS Ethics Compliance Survey
* Request for version addition

## Usage notes

* The notebook uses a service account for downloading content. You need the corresponding Nettskjema token to fetch submissions (this is not an EBRAINS token).
* The notebook is expected to be run in the EBRAINS Jupyter Lab environment, but you can also run it on your own machine if you have the token.
* `nettskjema_scripts.py` contains the necessary Python scripts for the notebook.

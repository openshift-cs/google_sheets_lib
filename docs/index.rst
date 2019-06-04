.. Google Sheets Library documentation master file, created by
   sphinx-quickstart on Tue Nov  6 18:02:02 2018.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to Google Sheets Library's documentation!
=================================================

.. toctree::
   :maxdepth: 3
   :caption: Contents:


Usage
=====

Create your `client credentials <https://cloud.google.com/docs/authentication/end-user#creating_your_client_credentials>`_
and place in your project root. Then create your Google Sheets client like::

   from google_sheets_lib import GoogleSheets

   client = GoogleSheets()

Alternatively, create a `service account <https://cloud.google.com/docs/authentication/production#obtaining_and_providing_service_account_credentials_manually>`_
and place the JSON in your project root. Then create your Google Sheets client like::

   from google_sheets_lib import GoogleSheets
   from pathlib import Path

   client = GoogleSheets(service_account_file=Path.cwd() / 'client_secret.json')


Change Log
==========

0.3
   * Added service account authentication

API
===

.. automodule:: google_sheets_lib
   :members:
   :undoc-members:
   :show-inheritance:

Index
=====

* :ref:`genindex`

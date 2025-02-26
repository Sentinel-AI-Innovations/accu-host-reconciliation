# ACCU - Host Asset System Reconciliation

This program is designed to reconcile the host asset system for the various ACCU host asset systems. It is built using Angular and the Angular CLI.

The program is designed to accept an Excel file that has tabs for Kace, AD, Blumira, Intune, Tenable, and Trend. The customer has determined the format of the Excel file, but in general there are columns for hostname, IP address, and sometimes other fields.

This application will use Kace as the "system of truth" and compare the other host asset systems against the information in the Kace tab, noting any discrepancies.

The customer requires certain handling of certain hostnames to suit their business, such as ignoring certain hostnames or treating them as special cases. This logic is included in the application.

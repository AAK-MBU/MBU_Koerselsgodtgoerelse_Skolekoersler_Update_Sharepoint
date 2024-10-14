## MBU Kørselsgodtgørelse for Skolekørsler - Sharepoint Updater

This robot is built for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).
It uploads the updated excel file to Sharepoint in the correct folder - Fejlet or Behandlet.
Deletes the old excel file in Til Udbetaling folder. 
Uploads attachements for failed elements.

### Arguments:

- **path**: The same path as the `path` argument in the uploader robot or the location where the Excel file is stored.

### Related robots:
    - Queue uploader.
    - Queue handler.
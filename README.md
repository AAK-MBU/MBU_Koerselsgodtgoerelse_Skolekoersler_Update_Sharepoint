## Kørselsgodtgørelse for Skolekørsler - Update Sharepoint

This robot is part of the 'MBU Koerselsgodtgoerelse Skolekoersler' process.

### Process Overview

The process consists of four main robots working in sequence:

1. **Create Excel and Upload to SharePoint**:  
   The first robot retrieves and exports weekly 'Egenbefordring' data from a database to an Excel file, which is then uploaded to SharePoint at the following location: `MBU - RPA - Egenbefordring/Dokumenter/General`. Once the file is processed, personnel will move it to `MBU - RPA - Egenbefordring/Dokumenter/General/Til udbetaling`. Run it with the 'Single Trigger' or with the Scheduled Trigger'.

2. **Queue Uploader**:  
   The second robot retrieves data from the Excel file and uploads it to the **Koerselsgodtgoerelse_egenbefordring** queue using [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator). Run it with the 'Single Trigger'.

3. **Queue Handler**:  
   The third robot, triggered by the Queue Trigger in OpenOrchestrator, processes the queue elements by creating tickets in OPUS.

4. **Update SharePoint (This robot)**:  
   The fourth robot cleans and updates the files in SharePoint by uploading the updated Excel file and attachments of any failed elements. Run it with the 'Single Trigger'.

### The Update Sharepoint Process

    - Upload the local Excel file to the Behandlet folder if no elements failed. If any failed elements it uploads it to the `Fejlet` folder
    - If any failed elements - the corresponding attachement gets uploaded to a folder with the same same as the Excel file inside the Fejlet folder.
    - Deletes the Excel file from the `Til Udbetaling folder`. 


### Process and Related Robots

1. **Create Excel & Upload to SharePoint**: [Create Excel & Upload To SharePoint](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Dan_Excel_Upload_Til_SharePoint)
2. **Queue Uploader** [Queue Uploader](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue_Uploader).
3. **Queue Handler**: [Queue Handler](https://github.com/AAK-MBU/MBU_Koerselsgodtgoerelse_Skolekoersler_Queue__Handler)
4. **Update SharePoint**: (This Robot)

### Arguments

- **path**: The same path as the `path` argument in the handler robot or the location where the Excel file is stored.

# import csv files to microsoft access

Public Sub ImportCSVFile(fileName as String)

    DoCmd.TransferText acImportDelim, "Specification file", "tblName", _
            fileName, true, , 1252

End Sub

# Convert a file from XLS to XLSX in order to be able to manipulate it

soffice --convert-to xlsx "/Users/alexandro666/Development/AlyCont/test.xls" -outdir "/Users/alexandro666/Development/AlyCont"

- Define where the actual values are in the XLSX are situated
  - For example, the total number of pages ( Pagina 1 din 33 )
  - Every field ( Data, Seria, Numar etc. )
- Start filtering. Eliminate the ones with 0 values as specified in the mail

# When creating the new XLSX file

- Have 2 entries, one for the 9% and one for 19% VAT ( if the specific VAT is applicable )

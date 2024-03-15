# Excel Processing

This project is for working with an excel file and processing it. It can read comments from the cells, choose the desired one and put it as the value of that cell. It can also count how many times a specific value is repeated in a specific column or row. 

---

- In the first part of the code, An excel file is read. There are some comments written in the comment section of each cell. The code reads the comments and randomly fill the value of the cell with one of those names in the comment. ***a name can not be repeated in one row.***
- In the second part of the code, you can keep count of how many times a specific name is repeated somewhere. It reads an input file which is a filled excel file and process that to find the counts. Then, the output which contains of the details is saved in a text file.

---

The libraries used in this project are: 

1. openpyxl (for working with excel files) 
2. pandas
3. random

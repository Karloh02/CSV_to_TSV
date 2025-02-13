The program was made under a assumption that the CSV file comes with a ";" sepparating the data, so it uses this information to organize the information. 
Since it is an specific use, there are some considerations regarding the starting line to read, and some manipulation of the header line. This is simple stuff, so you can adapt to your project with little effort. 

The program also changes the decimal operator "." and changes it to ",", also simple stuff that you can change in order to adapt to your project. 

I ended up using the aspose library to write the TSV file, since in my tests using only the rename function to .TSV did not work, but this can be caused due to my lack of coding skills 

If you end up reading this, and have any comments feel free to reach me out in my e-mail: gabrielraka@hotmail.com 

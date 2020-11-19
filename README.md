Hi its my first package for C#.
==================================

It's a small package for working with Excel files. Which simplifies the process of reading and creating excel files.


### 1. Read file.

For reading excel files package, has class Open with some methods.

Example : 
FullName        | prop1 | prop2 | prop3
----------------|-------|-------|------
Khamraev Akbar  | 1     | 2     | 0
Aliev Alisher   | 2     | 4     | 22
Valiev Vali     | 1     | 5     | -2

You have Excel file with this data and you have to save FullName,prop1,prop2,prop3 to your db.
***
With this package you can get data from file with 2 lines of code.

    1.Excel.Open excelFile = new Excel.Open(pathToFile); // For initialize class
    2.List<ExampleType> data =  excelFile.toList<ExampleType>(1); // For getting data

So how it's works?
***
In first line we create instance of Excel.Open class and give one argument, its path to file on your machine. Constructor also can take second (not required) argument isCurrentDirectory(default = true) which need to indicate what we works with file on current directory or not. 
***
In second line takes data from file. By using Excel.Open's method toList<T>()
  
    List<T> toList<T>(int row = 0, int colomn = 1,int WorksheetsIndex = 0)

Method takes e (not required) arguments :
* row(0) = need to set from what row you want to read
* colomn(1) = need to set from what colomn you want to read
* WorksheetsIndex(0) need to set from which workshet read data

Class also has method ParseTo<T> 
  
     T RowTo<T>(int row = 0, int colomn = 1,int WorksheetsIndex = 0)
Which parse one row to obj(T)

End indexer 

     string this[int row, int column]
What's return value on current cell




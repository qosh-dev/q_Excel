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

By the way class ExampleType looks like this 
    
    public class ExampleType
    {
        public string FullName { get; set; }
        public int prop1 { get; set; }
        public int prop2 { get; set; }
        public int prop3 { get; set; }
    }

***

Class also has method ParseTo<T> 
  
     T RowTo<T>(int row = 0, int colomn = 1,int WorksheetsIndex = 0)
Which parse one row to obj(T)

End indexer 

     string this[int row, int column]
What's return value of current cell



### 2. Create file
There are 3 was to create Excel file 


* byte[] Build(Action<ExcelPackage> result)
* void Build(Action<ExcelPackage> result, string path)
* FileContentResult Excel(this ControllerBase obj, Action<ExcelPackage> result,string fileName = "file")

To example I will use third method whats return file content on Asp.Net Core application
So lets create excel file with data like in example in top.

        public IActionResult Index()
        {
            List<ExampleType> list = new List<ExampleType>(...); 
            return  this.Excel(e => {
                var firstWorkSheet = e.addWorkSheet("workSheetName");
                firstWorkSheet.AddHeaders(headers);
                firstWorkSheet.AddLoop(list);
            });
        }
 Thats all you need to to do, or you can do more shotly

        public IActionResult Index()
        {
            List<string> headers = new List<sting>(){"FullName","prop1,"prop2","prop3"};
            List<ExampleType> list = new List<ExampleType>(...); 
            return this.Excel(e => e.addWorkSheet("workSheetName").AddHeaders(headers).AddLoop(list));
        }
        

Methods                   | prop1                           | args(required)                            | args(required)
--------------------------|---------------------------------|-------------------------------------------|---------------
indexer                   | set value to cell               | int row, int column                       |
indexer                   | set value to cells              | int row, int column, int row2, int column2|
AddHeaders                | Add bold text to cells          | T[] headersList                           | int colomn = 1, int row = 0, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Left, bool isBold = true, int fontSize = 12
AddMergedHeaders          | Add merged headers              | string[] valueIndexArr, int step          | int row = 0, int colomn = 0, ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Center, bool isBold = true, int fontSize = 12
AddLoop                   | Add your List to file           | List<T> list                              | int colomn = 1, int row = 0
AddLoopVertical           | Add your List to file(vertical) | List<T> list                              | int colomn = 1, int row = 0


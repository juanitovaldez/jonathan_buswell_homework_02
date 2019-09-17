# Unit 2 | Assignment - The VBA of Wall Street

The script summary.vbs makes some dangerous assumptions about the imutability of the data structure in the dataset. 
Got hit by a divide by zero buffer overload.
The performance is terrible. The nested for loop traverses the ticker column for each unique ticker. (797710x3169) = 2527942990 loops

Some thoughts to improve upon:
* It may be faster to read the columns into an array or a dictionary for vba to process.
* Perhaps instead of using VBA objects, I could just use vba to script some worksheet functions
* Break Apart the nested for loop
* Make an assumption about how the data is sorted to break
* Wrap up individual steps into functions to call
* Make the processes more generic and portable
* Use named ranges to save some hardcoding.

![2016 Stock Marker](Images/2016_stock.Png)
![2015 Stock Marker](Images/2015_stock.Png)
![2014 Stock Marker](Images/2014_stock.Png)
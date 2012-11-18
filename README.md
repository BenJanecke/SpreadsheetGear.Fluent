SpreadsheetGear.Fluent
======================

SpreadsheetGear.Fluent a fluent Api for the IRange interface. 

This Goal of this project is to allow developers to use a terse easy to understand Api when generating a spreadsheet with 
SpreadsheetGear or anything that implements IRange as defined by SpreadsheetGear.

Why?

Because I want developers to be able to do the following

	var ws = Factory.GetWorksheet();
	int rowStart = 1;
	int colStart = 1;
	int rowEnd = 2;
	int colEnd = 2;
	double? currency = 22;
	DateTime? Date = new DateTime(1,1,1);
	
	//Write A Currency
	ws.FluentCells(rowStart, colStart, rowEnd, colEnd).
		 SetValue(currency, NumberFormat.Currency).
		 SetStyle(style  => style.FontWeight = FontWeight.Bold).
		 ToggleMerge();

	rowStart++;
	colStart++;
	rowEnd++;
	colEnd++;
	
	//Write A Date 
	ws.FluentCells(rowStart, colStart, rowEnd, colEnd).
		 SetValue(date, NumberFormat.ShortDate).
		 SetStyle(style  => style.Font.Bold = True).
		 ToggleMerge();
	
	rowStart++;
	colStart++;
	rowEnd++;
	colEnd++;
	
	//Write with format 
	ws.FluentCells(rowStart, colStart, rowEnd, colEnd).
		 SetValueFormat("{0} Days Of Summer", 500).
		 SetStyle(style  => style.Font.Bold = True).
		 ToggleMerge();
  
	
	
Instead Of

	
	var ws = Factory.GetWorksheet();
	int rowStart = 1;
	int colStart = 1;
	int rowEnd = 2;
	int colEnd = 2;
	double? currency = 22;
	DateTime? Date = new DateTime(1,1,1);
	

	//Write A Currency
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Formula = currency;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].NumberFormat = NumberFormat.Currency;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Style.Font.Bold = True;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Merge();	 
			
	rowStart++;
	colStart++;
	rowEnd++;
	colEnd++;
	
	//Write A Date 
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Formula = date;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].NumberFormat = NumberFormat.ShortDate;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Style.Font.Bold = True;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Merge();	
	
	rowStart++;
	colStart++;
	rowEnd++;
	colEnd++;
	
	//Write with format 
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Formula = String.Format("{0} Days Of Summer", 500);
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Style.Font.Bold = True;
	ws.Cells[rowStart, colStart, rowEnd, colEnd].Merge();	

Disclaimer 

I dont work for or with SpreadhseetGear and I have not worked with them in the past.
Wich is to say I am in no way Affiliated with SpreadhseetGear LLC I just happen to really like their product.
This Api merely exists To Simplify my life and the lives of my developers.


Caveat 

SpreadsheetGear Is not free you have to buy it wich means I cannot upload or Refrence the SpreadsheetGear 
assemblies directly from within this project therefore whis Project stricly only communicates with the SpreadsheetGear 
Interfaces.

Imaginaaation 

Like I mentioned above a few times This Api Leverages The IRange interface not SpreadhseetGear directly. 

So How does that help you?

If SpreadsheetGear Is not your Component of choice simply write an adapter for your compnent to the IRange interface
And tada! you get to make use of this slick and sexy api.


clc;
clear all;
pkg load io;
pkg load windows;

sourceFileName = 'Failures Report.xlsx';
templateFileName = 'Template.xlsx';
reportFileName = 'RRRRRRRR.xlsx';



[filetype, sh_names, fformat, nmranges] = xlsfinfo(sourceFileName);
sheetNames = sh_names(:,1);
sheetNumber = size(sheetNames);


excelApp = actxserver ('Excel.Application');
workBookTemplate = excelApp.Workbooks.Open([(pwd()) '\' templateFileName]);
workBookFailures = excelApp.Workbooks.Open([(pwd()) '\' sourceFileName]);
workBookReport = excelApp.Workbooks.Add();
% excelApp.Visible = 1;

templateSheet = workBookTemplate.Sheets.Item('Name');
templateSheet.Activate();

for i = 1 : sheetNumber
	templateSheet.Copy(workBookReport.Sheets.Item(1));
	disp(i);
endfor

for i = 1 : sheetNumber
	tempVar = workBookReport.Worksheets.Item(i);
	tempVar.Name = sheetNames{i};
	disp(tempVar.Name);
endfor

for i = 1 : sheetNumber
	
	dataSheet = workBookFailures.Sheets.Item(sheetNames{i});
	reportSheet = workBookReport.Sheets.Item(sheetNames{i});

	dataSheet.Activate();
	workOrderData = dataSheet.Range('A2:A608');
	dateData = dataSheet.Range('H2:H608');
	qtyData = dataSheet.Range('AW2:AW608');
	% workOrderData = dataSheet.Range('A2:A608');
	% dateData = dataSheet.Range('B2:B608');
	% qtyData = dataSheet.Range('C2:C608');

	reportSheet.Activate();
	workOrderCell = reportSheet.Range('A2:A608');
	dateCell = reportSheet.Range('B2:B608');
	qtyCell = reportSheet.Range('C2:C608');
	workOrderCell.Value = workOrderData.Value;
	dateCell.Value = dateData.Value;
	qtyCell.Value = qtyData.Value; 

endfor

workBookReport.Sheets.Item('Sheet1').Delete;
workBookReport.Activate();
excelApp.ActiveWorkbook.SaveAs([(pwd()) '\' reportFileName]);

excelApp.DisplayAlerts = false;
excelApp.Quit();



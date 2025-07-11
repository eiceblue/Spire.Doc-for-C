#include "pch.h"

using namespace Spire::Doc;

void setTableBorders(intrusive_ptr<Table> table)
{
	table->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetFormat()->GetBorders()->SetLineWidth(3.0F);
	table->GetFormat()->GetBorders()->SetColor(Color::GetRed());
}

void setCellBorders(intrusive_ptr<TableCell> tableCell)
{
	tableCell->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::DotDash);
	tableCell->GetCellFormat()->GetBorders()->SetLineWidth(1.0F);
	tableCell->GetCellFormat()->GetBorders()->SetColor(Color::GetGreen());
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/DifferentBorders.docx";
	std::wstring inputFile = DATAPATH"/TableSample.docx";
	

	//Open a Word document as template
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(document->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

	//Set borders of table
	setTableBorders(table);

	//Set borders of cell
	setCellBorders(table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0));

	//Save and launch document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
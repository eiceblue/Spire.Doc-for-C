#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateTableDirectly.docx";
	
	//Create a Word document
	intrusive_ptr<Document> doc = new Document();

	//Add a section
	intrusive_ptr<Section> section = doc->AddSection();

	//Create a table 
	intrusive_ptr<Table> table = new Table(doc);
	table->ResetCells(1, 2);
	//Set the width of table
	table->SetPreferredWidth(new PreferredWidth(WidthType::Percentage, 100));
	//Set the border of table
	table->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);

	//Create a table row
	intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(0);
	row->SetHeight(50.0f);

	//Create a table cell
	intrusive_ptr<TableCell> cell1 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);
	//Add a paragraph
	intrusive_ptr<Paragraph> para1 = cell1->AddParagraph();
	//Append text in the paragraph
	para1->AppendText(L"Row 1, Cell 1");
	//Set the horizontal alignment of paragrah
	para1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	//Set the background color of cell
	cell1->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());
	//Set the vertical alignment of paragraph
	cell1->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);

	//Create a table cell
	intrusive_ptr<TableCell> cell2 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1);
	intrusive_ptr<Paragraph> para2 = cell2->AddParagraph();
	para2->AppendText(L"Row 1, Cell 2");
	para2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	cell2->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());
	cell2->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	row->GetCells()->Add(cell2);

	//Add the table in the section
	section->GetTables()->Add(table);

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();

}

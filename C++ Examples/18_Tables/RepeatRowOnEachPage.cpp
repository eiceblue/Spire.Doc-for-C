#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RepeatRowOnEachPage.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	//Create a new section
	intrusive_ptr<Section> section = document->AddSection();

	//Create a table width default borders
	intrusive_ptr<Table> table = section->AddTable(true);
	//Set table with to 100%
	intrusive_ptr<PreferredWidth> width = new PreferredWidth(WidthType::Percentage, 100);
	table->SetPreferredWidth(width);

	//Add a new row 
	intrusive_ptr<TableRow> row = table->AddRow();
	//Set the row as a table header 
	row->SetIsHeader(true);

	//Add a new cell for row
	intrusive_ptr<TableCell> cell = row->AddCell();

	//Set the backcolor of row
	for (int i = 0; i < row->GetCells()->GetCount(); i++)
	{
		row->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetLightGray());
	}

	cell->SetCellWidth(100, CellWidthType::Percentage);
	//Add a paragraph for cell to put some data
	intrusive_ptr<Paragraph> parapraph = cell->AddParagraph();
	//Add text 
	parapraph->AppendText(L"Row Header 1");
	//Set paragraph horizontal center alignment
	parapraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	row = table->AddRow(false, 1);
	row->SetIsHeader(true);

	for (int i = 0; i < row->GetCells()->GetCount(); i++)
	{
		row->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetIvory());
	}

	//Set row height
	row->SetHeight(30);
	cell = row->GetCells()->GetItemInCellCollection(0);
	cell->SetCellWidth(100, CellWidthType::Percentage);
	//Set cell vertical middle alignment
	cell->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	//Add a paragraph for cell to put some data
	parapraph = cell->AddParagraph();
	//Add text 
	parapraph->AppendText(L"Row Header 2");
	parapraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Add many common rows 
	for (int i = 0; i < 70; i++)
	{
		row = table->AddRow(false, 2);
		cell = row->GetCells()->GetItemInCellCollection(0);
		//Set cell width to 50% of table width
		cell->SetCellWidth(50, CellWidthType::Percentage);
		cell->AddParagraph()->AppendText(L"Column 1 Text");
		cell = row->GetCells()->GetItemInCellCollection(1);
		cell->SetCellWidth(50, CellWidthType::Percentage);
		cell->AddParagraph()->AppendText(L"Column 2 Text");
	}
	//Set cell backcolor
	for (int j = 1; j < table->GetRows()->GetCount(); j++)
	{
		if (j % 2 == 0)
		{
			intrusive_ptr<TableRow> row2 = table->GetRows()->GetItemInRowCollection(j);
			for (int f = 0; f < row2->GetCells()->GetCount(); f++)
			{
				row2->GetCells()->GetItemInCellCollection(f)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetLightBlue());
			}
		}

	}

	//Save file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}

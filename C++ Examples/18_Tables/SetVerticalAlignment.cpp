#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetVerticalAlignment.docx";
	std::wstring inputFile = DATAPATH"/E-iceblue.png";

	//Create a new Word document and add a new section
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> section = doc->AddSection();

	//Add a table with 3 columns and 3 rows
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(3, 3);

	//Merge rows
	table->ApplyVerticalMerge(0, 0, 2);

	//Set the vertical alignment for each cell, default is top
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);
	table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);

	//Inset a picture into the table cell
	intrusive_ptr<Paragraph> paraPic = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph();

	intrusive_ptr<DocPicture> pic = paraPic->AppendPicture(inputFile.c_str());

	//Create data
	std::vector<std::vector<std::wstring>> data =
	{
		{L"", L"Spire.Office", L"Spire.DataExport"},
		{L"", L"Spire.Doc", L"Spire.DocViewer"},
		{L"", L"Spire.XLS", L"Spire.PDF"}
	};

	//Append data to table
	for (int r = 0; r < 3; r++)
	{
		intrusive_ptr<TableRow> dataRow = table->GetRows()->GetItemInRowCollection(r);
		dataRow->SetHeight(50);
		for (int c = 0; c < 3; c++)
		{
			if (c == 1)
			{
				intrusive_ptr<Paragraph> par = dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				dataRow->GetCells()->GetItemInCellCollection(c)->SetCellWidth((section->GetPageSetup()->GetClientWidth()) / 2, CellWidthType::Point);
			}
			if (c == 2)
			{
				intrusive_ptr<Paragraph> par = dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				dataRow->GetCells()->GetItemInCellCollection(c)->SetCellWidth((section->GetPageSetup()->GetClientWidth()) / 2, CellWidthType::Point);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}

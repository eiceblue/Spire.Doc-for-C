#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateVerticalTable.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a table with rows and columns and set the text for the table.
	intrusive_ptr<Table> table = section->AddTable();
	table->ResetCells(1, 1);
	intrusive_ptr<TableCell> cell = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);
	table->GetRows()->GetItemInRowCollection(0)->SetHeight(150);
	cell->AddParagraph()->AppendText(L"Draft copy in vertical style");

	//Set the TextDirection for the table to RightToLeftRotated.
	cell->GetCellFormat()->SetTextDirection(TextDirection::RightToLeftRotated);

	//Set the table format.
	table->GetFormat()->SetWrapTextAround(true);
	table->GetFormat()->GetPositioning()->SetVertRelationTo(VerticalRelation::Page);
	table->GetFormat()->GetPositioning()->SetHorizRelationTo(HorizontalRelation::Page);
	table->GetFormat()->GetPositioning()->SetHorizPosition(section->GetPageSetup()->GetPageSize()->GetWidth() - table->GetWidth());
	table->GetFormat()->GetPositioning()->SetVertPosition(200);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}


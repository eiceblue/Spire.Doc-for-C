#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateNestedTable.docx";

	//Create a new document
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> section = doc->AddSection();

	//Add a table
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(2, 2);

	//Set column width
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->AutoFit(AutoFitBehaviorType::AutoFitToWindow);

	//Insert content to cells
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"Spire.Doc for .NET");
	std::wstring text = L"Spire.Doc for .NET is a professional Word"
		L".NET library specifically designed for developers to create,"
		L"read, write, convert and print Word document files from any .NET"
		L"platform with fast and high quality performance.";
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(text.c_str());

	//Add a nested table to cell(first row, second column)
	intrusive_ptr<Table> nestedTable = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddTable(true);
	nestedTable->ResetCells(4, 3);
	nestedTable->AutoFit(AutoFitBehaviorType::AutoFitToContents);

	//Add content to nested cells
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"NO.");
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Item");
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"Price");

	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"1");
	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Pro Edition");
	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$799");

	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"2");
	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Standard Edition");
	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$599");

	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"3");
	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Free Edition");
	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$0");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

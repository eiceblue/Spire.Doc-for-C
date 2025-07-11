#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CloneTable.docx";
	std::wstring inputFile = DATAPATH"/TableTemplate.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> se = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table
	intrusive_ptr<Table> original_Table = Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0));

	//Copy the existing table to copied_Table via Table.clone()
	intrusive_ptr<Table> copied_Table = original_Table->CloneTable();
	std::vector<std::wstring> st = { L"Spire.Presentation for .Net", L"A professional PowerPoint® compatible library that enables developers to create, read, write, modify, convert and Print PowerPoint documents on any .NET framework, .NET Core platform." };
	//Get the last row of table
	intrusive_ptr<TableRow> lastRow = copied_Table->GetRows()->GetItemInRowCollection(copied_Table->GetRows()->GetCount() - 1);
	//Change last row data
	for (int i = 0; i < lastRow->GetCells()->GetCount() - 1; i++)
	{
		lastRow->GetCells()->GetItemInCellCollection(i)->GetParagraphs()->GetItemInParagraphCollection(0)->SetText(st[i].c_str());
	}
	//Add copied_Table in section
	se->GetTables()->Add(copied_Table);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

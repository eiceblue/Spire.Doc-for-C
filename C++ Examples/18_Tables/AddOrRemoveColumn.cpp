#include "pch.h"

using namespace Spire::Doc;

void AddColumn(intrusive_ptr<Table> table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		intrusive_ptr<TableCell> addCell = new TableCell(table->GetDocument());
		table->GetRows()->GetItemInRowCollection(r)->GetCells()->Insert(columnIndex, addCell);
	}
}

void RemoveColumn(intrusive_ptr<Table> table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		table->GetRows()->GetItemInRowCollection(r)->GetCells()->RemoveAt(columnIndex);
	}
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddOrRemoveColumn.docx";
	std::wstring inputFile = DATAPATH"/Template_N2.docx";

	//Load the document from disk.
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Access the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Access the first table
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Add a blank column
	int columnIndex1 = 0;
	AddColumn(table, columnIndex1);

	//Remove a column
	int columnIndex2 = 2;
	RemoveColumn(table, columnIndex2);

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
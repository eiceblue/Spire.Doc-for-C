#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SplitTable.docx";
	std::wstring inputFile = DATAPATH"/CombineAndSplitTables.docx";

	//Load document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//We will split the table at the third row;
	int splitIndex = 2;

	//Create a new table for the split table
	intrusive_ptr<Table> newTable = new Table(section->GetDocument());

	//Add rows to the new table
	for (int i = splitIndex; i < table->GetRows()->GetCount(); i++)
	{
		newTable->GetRows()->Add(table->GetRows()->GetItemInRowCollection(i)->CloneTableRow());
	}

	//Remove rows from original table
	for (int i = table->GetRows()->GetCount() - 1; i >= splitIndex; i--)
	{
		table->GetRows()->RemoveAt(i);
	}

	//Add the new table in section
	section->GetTables()->Add(newTable);

	//Save the Word file
	section->GetDocument()->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CloneRow.docx";
	std::wstring inputFile = DATAPATH"/TableTemplate.docx";

	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> se = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first row of the first table
	intrusive_ptr<TableRow> firstRow = Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->GetItemInRowCollection(0);

	//Copy the first row to clone_FirstRow via TableRow.clone()
	intrusive_ptr<TableRow> clone_FirstRow = firstRow->CloneTableRow();

	Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->Add(clone_FirstRow);
	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

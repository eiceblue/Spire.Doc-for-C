#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetColumnWidth.docx";
	std::wstring inputFile = DATAPATH"/TableSample.docx";

	//Create a document and load file
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Traverse the first column
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		//Set the width and type of the cell
		table->GetRows()->GetItemInRowCollection(i)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(200, CellWidthType::Point);
	}

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
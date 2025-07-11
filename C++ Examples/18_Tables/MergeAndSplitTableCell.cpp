#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/MergeAndSplitTableCell.docx";
	std::wstring inputFile = DATAPATH"/TableSample.docx";

	//Create a document and load file from disk
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));
	//The method shows how to merge cell horizontally
	table->ApplyHorizontalMerge(6, 2, 3);
	//The method shows how to merge cell vertically
	table->ApplyVerticalMerge(2, 4, 5);
	//The method shows how to split the cell
	table->GetRows()->GetItemInRowCollection(8)->GetCells()->GetItemInCellCollection(3)->SplitCell(2, 2);
	//Save to file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
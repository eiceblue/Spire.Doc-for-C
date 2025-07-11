#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveTable.docx";
	std::wstring inputFile = DATAPATH"/Template.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Remove the first Table            
	doc->GetSections()->GetItemInSectionCollection(0)->GetTables()->RemoveAt(0);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

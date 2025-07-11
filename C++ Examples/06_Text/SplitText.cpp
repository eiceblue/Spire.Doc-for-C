#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"SplitText.docx";

	//Create a new document and load from file
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Add a column to the first section and set width and spacing
	doc->GetSections()->GetItemInSectionCollection(0)->AddColumn(100.0f, 20.0f);
	//Add a line between the two columns
	doc->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetColumnsLineBetween(true);

	//Save and launch the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}


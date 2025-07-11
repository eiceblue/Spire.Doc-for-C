#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AdjustHeaderFooterHeight.docx";
	std::wstring inputFile = DATAPATH"/HeaderAndFooter.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Adjust the height of headers in the section
	section->GetPageSetup()->SetHeaderDistance(100);

	//Adjust the height of footers in the section
	section->GetPageSetup()->SetFooterDistance(100);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
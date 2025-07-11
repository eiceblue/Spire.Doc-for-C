#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DifferentPageSetup.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DifferentPageSetup.docx";


	//Open a Word document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());


	//Get the second section 
	intrusive_ptr<Section> SectionTwo = doc->GetSections()->GetItemInSectionCollection(1);

	//Set the orientation
	SectionTwo->GetPageSetup()->SetOrientation(PageOrientation::Landscape);

	//Set page size
	//SectionTwo.PageSetup.PageSize = new SizeF(800, 800);

	doc->SaveToFile(outputFile.c_str());
	doc->Close();

}

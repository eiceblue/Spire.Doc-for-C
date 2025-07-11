#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_N2.docx";
	wstring outputFile = output_path + L"ModifyPageSetupOfSection.docx";

	//Load Word from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Loop through all sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
		//Modify the margins
		section->GetPageSetup()->SetMargins(new MarginsF(100, 80, 100, 80));
		//Modify the page size
		section->GetPageSetup()->SetPageSize(PageSize::Letter());
	}

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
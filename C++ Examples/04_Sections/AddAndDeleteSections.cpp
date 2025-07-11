#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SectionTemplate.docx";
	wstring outputFile = output_path + L"AddAndDeleteSections.docx";

	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Add a section
	doc->AddSection();
	//Delete the last section
	doc->GetSections()->RemoveAt(doc->GetSections()->GetCount() - 1);

	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();

}


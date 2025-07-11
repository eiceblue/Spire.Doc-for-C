#include "pch.h"



using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SectionTemplate.docx";
	wstring outputFile = output_path + L"CloneSection.docx";
	//Load source file
	intrusive_ptr<Document> srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile.c_str());

	//Create destination file
	intrusive_ptr<Document> desDoc = new Document();

	intrusive_ptr<Section> cloneSection = nullptr;
	for (int i = 0; i < srcDoc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = srcDoc->GetSections()->GetItemInSectionCollection(i);
		//Clone section
		cloneSection = section->CloneSection();
		//Add the cloneSection in destination file
		desDoc->GetSections()->Add(cloneSection);
	}
	//Save the Word
	desDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	desDoc->Close();
}

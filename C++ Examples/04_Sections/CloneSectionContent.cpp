#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SectionTemplate.docx";
	wstring outputFile = output_path + L"CloneSectionContent.docx";

	//Load the Word document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> sec1 = doc->GetSections()->GetItemInSectionCollection(0);
	//Get the second section
	intrusive_ptr<Section> sec2 = doc->GetSections()->GetItemInSectionCollection(1);

	//Loop through the contents of sec1
	for (int i = 0; i < sec1->GetBody()->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = sec1->GetBody()->GetChildObjects()->GetItem(i);
		//Clone the contents to sec2
		sec2->GetBody()->GetChildObjects()->Add(obj->Clone());
	}

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}

#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CopyHeaderAndFooter.docx";
	std::wstring inputFile = DATAPATH"/HeaderAndFooter.docx";
	std::wstring inputFile_1 = DATAPATH"/Template.docx";

	//Load the source file
	intrusive_ptr<Document> doc1 = new Document();
	doc1->LoadFromFile(inputFile.c_str());

	//Get the header section from the source document
	intrusive_ptr<HeaderFooter> header = doc1->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

	//Load the destination file
	intrusive_ptr<Document> doc2 = new Document();
	doc2->LoadFromFile(inputFile_1.c_str());

	//Copy each object in the header of source file to destination file
	for (int i = 0; i < doc2->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc2->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < header->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = header->GetChildObjects()->GetItem(j);
			section->GetHeadersFooters()->GetHeader()->GetChildObjects()->Add(obj->Clone());
		}
	}

	//Save and launch document
	doc2->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc1->Close();
	doc2->Close();
}
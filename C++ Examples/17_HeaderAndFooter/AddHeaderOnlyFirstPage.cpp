#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddHeaderOnlyFirstPage.docx";
	std::wstring inputFile = DATAPATH"/MultiplePages.docx";
	std::wstring inputFile_1 = DATAPATH"/HeaderAndFooter.docx";

	//Load the source file
	intrusive_ptr<Document> doc1 = new Document();
	doc1->LoadFromFile(inputFile_1.c_str());

	//Get the header from the first section
	intrusive_ptr<HeaderFooter> header = doc1->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

	//Load the destination file
	intrusive_ptr<Document> doc2 = new Document();
	doc2->LoadFromFile(inputFile.c_str());

	//Get the first page header of the destination document
	intrusive_ptr<HeaderFooter> firstPageHeader = doc2->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFirstPageHeader();

	//Specify that the current section has a different header/footer for the first page
	for (int i = 0; i < doc2->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = doc2->GetSections()->GetItemInSectionCollection(i);
		section->GetPageSetup()->SetDifferentFirstPageHeaderFooter(true);
	}

	//Removes all child objects in firstPageHeader
	firstPageHeader->GetParagraphs()->Clear();

	//Add all child objects of the header to firstPageHeader
	for (int j = 0; j < header->GetChildObjects()->GetCount(); j++)
	{
		intrusive_ptr<DocumentObject> obj = header->GetChildObjects()->GetItem(j);
		firstPageHeader->GetChildObjects()->Add(obj->Clone());
	}

	//Save and launch the file
	doc2->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc1->Close();
	doc2->Close();
}
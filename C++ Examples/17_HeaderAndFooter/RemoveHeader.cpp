#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveHeader.docx";
	std::wstring inputFile = DATAPATH"/HeaderAndFooter.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section of the document
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Traverse the word document and clear all headers in different type
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
			//Clear header in the first page
			intrusive_ptr<HeaderFooter> header;
			header = section->GetHeadersFooters()->GetFirstPageHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
			//Clear header in the odd page
			header = section->GetHeadersFooters()->GetOddHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
			//Clear header in the even page
			header = section->GetHeadersFooters()->GetEvenHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
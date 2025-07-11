#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveFooter.docx";
	std::wstring inputFile = DATAPATH"/HeaderAndFooter.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Traverse the word document and clear all footers in different type
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
			//Clear footer in the first page
			intrusive_ptr<HeaderFooter> footer;
			footer = section->GetHeadersFooters()->GetFirstPageFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
			//Clear footer in the odd page
			footer = section->GetHeadersFooters()->GetOddFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
			//Clear footer in the even page
			footer = section->GetHeadersFooters()->GetEvenFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
		}
	}


	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
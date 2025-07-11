#include "pch.h"

using namespace Spire::Doc;

void ExtractBetweenParagraphStyles(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, const std::wstring& stylename1, const std::wstring& stylename2)
{
	int startindex = 0;
	int endindex = 0;
	//travel the sections of source document

	for (int i = 0; i < sourceDocument->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = sourceDocument->GetSections()->GetItemInSectionCollection(i);
		//travel the paragraphs
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			//Judge paragraph style1
			if (paragraph->GetStyleName() == stylename1)
			{
				//Get the paragraph index
				startindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
			//Judge paragraph style2
			if (paragraph->GetStyleName() == stylename2)
			{
				//Get the paragraph index
				endindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
		}
		//Extract the content
		for (int i = startindex + 1; i < endindex; i++)
		{
			//Clone the ChildObjects of source document
			intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

			//Add to destination document 
			destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(doobj);
		}
	}
}

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"BetweenParagraphStyle.docx";
	wstring outputFile = output_path + L"BetweenParagraphStyle.docx";

	//Create the first document
	intrusive_ptr<Document> sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add a section
	intrusive_ptr<Section> section = destinationDoc->AddSection();

	//Extract content between the first paragraph to the third paragraph
	ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, L"1", L"2");

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();

}

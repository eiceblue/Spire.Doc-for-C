#include "pch.h"
using namespace Spire::Doc;

int main() 
{
	std::wstring outputFile = OUTPUTPATH"/ReplaceImageWithText.docx";
	std::wstring inputFile = DATAPATH"/ImageTemplate.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Replace all pictures with texts
	int j = 1;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
		for (int k = 0; k < sec->GetParagraphs()->GetCount(); k++)
		{
			intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(k);
			std::vector<intrusive_ptr<DocumentObject>> pictures;
			//Get all pictures in the Word document
			for (int m = 0; m < para->GetChildObjects()->GetCount(); m++)
			{
				intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(m);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}

			//Replace pitures with the text "Here was image {image index}"
			for (auto pic : pictures)
			{
				int index = para->GetChildObjects()->IndexOf(pic);
				intrusive_ptr<TextRange> range = new TextRange(doc);
				wstring temp = L"Here was image " + to_wstring(j) + L"";
				range->SetText(temp.c_str());
				para->GetChildObjects()->Insert(index, range);
				para->GetChildObjects()->Remove(pic);
				j++;
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
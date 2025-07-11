#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/UpdateImage.docx";
	std::wstring inputFile = DATAPATH"/ImageTemplate.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all pictures in the Word document
	std::vector<intrusive_ptr<DocumentObject>> pictures;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(k);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}
		}
	}

	//Replace the first picture with a new image file
	intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(pictures[0]);
	picture->LoadImageSpire(DATAPATH"/E-iceblue.png");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
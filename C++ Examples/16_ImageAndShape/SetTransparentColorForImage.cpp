#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SetTransparentColorForImage.docx";
	std::wstring inputFile = DATAPATH"/ImageTemplate.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first paragraph in the first section
	intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

	//Set the blue color of the image(s) in the paragraph to transperant
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(i);
		if (Object::CheckType<DocPicture>(obj))
		{
			intrusive_ptr<DocPicture> picture = boost::dynamic_pointer_cast<DocPicture>(obj);
			picture->SetTransparentColor(Color::GetBlue());
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
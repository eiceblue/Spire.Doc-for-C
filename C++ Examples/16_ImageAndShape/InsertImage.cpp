#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/InsertImage.docx";
	std::wstring inputFile = DATAPATH"/BlankTemplate.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();
	paragraph->AppendText(L"The sample demonstrates how to insert an image into a document.");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"The above is a picture.");
	//Create a picture
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	picture->LoadImageSpire(DATAPATH"/Word.png");
	//set image's position
	picture->SetHorizontalPosition(50.0F);
	picture->SetVerticalPosition(60.0F);

	//set image's size
	picture->SetWidth(200);
	picture->SetHeight(200);

	//set textWrappingStyle with image;
	picture->SetTextWrappingStyle(TextWrappingStyle::Through);
	//Insert the picture at the beginning of the second paragraph
	paragraph->GetChildObjects()->Insert(0, picture);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
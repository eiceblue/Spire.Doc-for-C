#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_5.docx";
	wstring outputFile = output_path + L"CopyParagraph.docx";

	//Create Word document1.
	intrusive_ptr<Document> document1 = new Document();

	//Load the file from disk.
	document1->LoadFromFile(inputFile.c_str());

	//Create a new document.
	intrusive_ptr<Document> document2 = new Document();

	//Get paragraph 1 and paragraph 2 in document1.
	intrusive_ptr<Section> s = document1->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> p1 = s->GetParagraphs()->GetItemInParagraphCollection(0);
	intrusive_ptr<Paragraph> p2 = s->GetParagraphs()->GetItemInParagraphCollection(1);

	//Copy p1 and p2 to document2.
	intrusive_ptr<Section> s2 = document2->AddSection();
	intrusive_ptr<Paragraph> NewPara1 = Object::Dynamic_cast<Paragraph>(p1->Clone());
	s2->GetParagraphs()->Add(NewPara1);
	intrusive_ptr<Paragraph> NewPara2 = Object::Dynamic_cast<Paragraph>(p2->Clone());
	s2->GetParagraphs()->Add(NewPara2);

	//Add watermark.
	intrusive_ptr<PictureWatermark> WM = new PictureWatermark();
	WM->SetPicture(DATAPATH"/Logo.jpg");

	document2->SetWatermark(WM);

	//Save the file.
	document2->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document2->Close();
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddPictureCaption.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Create a new section
	intrusive_ptr<Section> section  = document->AddSection();

	//Add the first picture
	intrusive_ptr<Paragraph> par1 = section->AddParagraph();
	par1->GetFormat()->SetAfterSpacing(10);
	intrusive_ptr<DocPicture> pic1 = par1->AppendPicture((input_path + L"Spire.Doc.png").c_str());
	pic1->SetHeight(100);
	pic1->SetWidth(120);
	//Add caption to the picture
	pic1->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

	//Add the second picture
	intrusive_ptr<Paragraph> par2 = section->AddParagraph();
	intrusive_ptr<DocPicture> pic2 = par2->AppendPicture((input_path + L"Word.png").c_str());
	pic2->SetHeight(100);
	pic2->SetWidth(120);
	//Add caption to the picture
	pic2->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

	//Update fields
	document->SetIsUpdateFields(true);

	//Save the Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
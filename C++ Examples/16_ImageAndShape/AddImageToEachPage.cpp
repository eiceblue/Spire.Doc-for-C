#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddImageToEachPage.docx";
	std::wstring inputFile = DATAPATH"/SampleB_2.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	std::wstring imgPath = DATAPATH"/Spire.Doc.png";

	//Add a picture in footer and set it's position

	intrusive_ptr<DocPicture> picture = document->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendPicture(imgPath.c_str());
	picture->SetVerticalOrigin(VerticalOrigin::Page);
	picture->SetHorizontalOrigin(HorizontalOrigin::Page);
	picture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);
	picture->SetTextWrappingStyle(TextWrappingStyle::None);

	//Add a textbox in footer and set it's positiion
	intrusive_ptr<TextBox> textbox = document->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendTextBox(150, 20);
	textbox->SetVerticalOrigin(VerticalOrigin::Page);
	textbox->SetHorizontalOrigin(HorizontalOrigin::Page);
	textbox->SetHorizontalPosition(300);
	textbox->SetVerticalPosition(700);
	textbox->GetBody()->AddParagraph()->AppendText(L"Welcome to E-iceblue");

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
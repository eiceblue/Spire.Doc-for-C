#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ImageHeaderAndFooter.docx";
	std::wstring inputFile = DATAPATH"/Template.docx";

	//Load the document from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the header of the first page
	intrusive_ptr<HeaderFooter> header = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

	//Add a paragraph for the header
	intrusive_ptr<Paragraph> paragraph = header->AddParagraph();

	//Set the format of the paragraph
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//Append a picture in the paragraph

	intrusive_ptr<DocPicture> headerimage = paragraph->AppendPicture(DATAPATH"/E-iceblue.png");

	headerimage->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//Get the footer of the first section
	intrusive_ptr<HeaderFooter> footer = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter();

	//Add a paragraph for the footer
	intrusive_ptr<Paragraph> paragraph2 = footer->AddParagraph();

	//Set the format of the paragraph
	paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	//Append a picture in the paragraph
	intrusive_ptr<DocPicture> footerimage = paragraph2->AppendPicture(DATAPATH"/logo.png");

	//Append text in the paragraph
	wstring string = L"Copyright \u00A9 2013 e-iceblue. All Rights Reserved.";
	intrusive_ptr<TextRange> TR = paragraph2->AppendText(string.c_str());
	TR->GetCharacterFormat()->SetFontName(L"Arial");
	TR->GetCharacterFormat()->SetFontSize(10);
	TR->GetCharacterFormat()->SetTextColor(Color::GetBlack());

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
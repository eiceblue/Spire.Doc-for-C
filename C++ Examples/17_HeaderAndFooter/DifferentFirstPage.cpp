#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/DifferentFirstPage.docx";
	std::wstring inputFile = DATAPATH"/MultiplePages.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the section and set the property true
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	section->GetPageSetup()->SetDifferentFirstPageHeaderFooter(true);

	//Set the first page header. Here we append a picture in the header
	intrusive_ptr<Paragraph> paragraph1 = section->GetHeadersFooters()->GetFirstPageHeader()->AddParagraph();
	paragraph1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	intrusive_ptr<DocPicture> headerimage = paragraph1->AppendPicture(DATAPATH"/E-iceblue.png");

	//Set the first page footer
	intrusive_ptr<Paragraph> paragraph2 = section->GetHeadersFooters()->GetFirstPageFooter()->AddParagraph();
	paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	intrusive_ptr<TextRange> FF = paragraph2->AppendText(L"First Page Footer");
	FF->GetCharacterFormat()->SetFontSize(10);

	//Set the other header & footer. If you only need the first page header & footer, don't set this
	intrusive_ptr<Paragraph> paragraph3 = section->GetHeadersFooters()->GetHeader()->AddParagraph();
	paragraph3->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	intrusive_ptr<TextRange> NH = paragraph3->AppendText(L"Spire.Doc for .NET");
	NH->GetCharacterFormat()->SetFontSize(10);

	intrusive_ptr<Paragraph> paragraph4 = section->GetHeadersFooters()->GetFooter()->AddParagraph();
	paragraph4->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	intrusive_ptr<TextRange> NF = paragraph4->AppendText(L"E-iceblue");
	NF->GetCharacterFormat()->SetFontSize(10);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}
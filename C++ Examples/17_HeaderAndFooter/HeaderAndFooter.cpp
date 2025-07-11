#include "pch.h"
using namespace Spire::Doc;

void InsertHeaderAndFooter(intrusive_ptr<Section> section)
{
	intrusive_ptr<HeaderFooter> header = section->GetHeadersFooters()->GetHeader();
	intrusive_ptr<HeaderFooter> footer = section->GetHeadersFooters()->GetFooter();

	//insert picture and text to header
	intrusive_ptr<Paragraph> headerParagraph = header->AddParagraph();

	intrusive_ptr<DocPicture> headerPicture = headerParagraph->AppendPicture(DATAPATH"/Header.png");

	//header text
	intrusive_ptr<TextRange> text = headerParagraph->AppendText(L"Demo of Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(10);
	text->GetCharacterFormat()->SetItalic(true);
	headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetSpace(0.05F);


	//header picture layout - text wrapping
	headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);

	//header picture layout - position
	headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

	//insert picture to footer
	intrusive_ptr<Paragraph> footerParagraph = footer->AddParagraph();

	intrusive_ptr<DocPicture> footerPicture = footerParagraph->AppendPicture(DATAPATH"/Footer.png");

	//footer picture layout
	footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
	footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//insert page number
	footerParagraph->AppendField(L"page number", FieldType::FieldPage);
	footerParagraph->AppendText(L" of ");
	footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
	footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetSpace(0.05F);
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/HeaderAndFooter.docx";
	std::wstring inputFile = DATAPATH"/Sample.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//insert header and footer
	InsertHeaderAndFooter(section);

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
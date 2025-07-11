#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"InsertEndnote.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertEndnote.docx";

	//Create a document and load file
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> s = doc->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> p = s->GetParagraphs()->GetItemInParagraphCollection(1);

	//add endnote
	intrusive_ptr<Footnote> endnote = p->AppendFootnote(FootnoteType::Endnote);

	//append text
	intrusive_ptr<TextRange> text = endnote->GetTextBody()->AddParagraph()->AppendText(L"Reference: Wikipedia");

	//set text format
	text->GetCharacterFormat()->SetFontName(L"Impact");
	text->GetCharacterFormat()->SetFontSize(14);
	text->GetCharacterFormat()->SetTextColor(Color::GetDarkOrange());

	//Set marker format of endnote
	endnote->GetMarkerCharacterFormat()->SetFontName(L"Calibri");
	endnote->GetMarkerCharacterFormat()->SetFontSize(25);
	endnote->GetMarkerCharacterFormat()->SetTextColor(Color::GetDarkBlue());

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}

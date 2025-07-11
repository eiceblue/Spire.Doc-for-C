#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FootnoteExample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertFootnote.docx";

	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<TextSelection> selection = document->FindString(L"Spire.Doc", false, true);

	intrusive_ptr<TextRange> textRange = selection->GetAsOneRange();
	intrusive_ptr<Paragraph> paragraph = textRange->GetOwnerParagraph();
	int index = paragraph->GetChildObjects()->IndexOf(textRange);
	intrusive_ptr<Footnote> footnote = paragraph->AppendFootnote(FootnoteType::Footnote);
	paragraph->GetChildObjects()->Insert(index + 1, footnote);

	textRange = footnote->GetTextBody()->AddParagraph()->AppendText(L"Welcome to evaluate Spire.Doc");
	textRange->GetCharacterFormat()->SetFontName(L"Arial Black");
	textRange->GetCharacterFormat()->SetFontSize(10);
	textRange->GetCharacterFormat()->SetTextColor(Color::GetDarkGray());

	footnote->GetMarkerCharacterFormat()->SetFontName(L"Calibri");
	footnote->GetMarkerCharacterFormat()->SetFontSize(12);
	footnote->GetMarkerCharacterFormat()->SetBold(true);
	footnote->GetMarkerCharacterFormat()->SetTextColor(Color::GetDarkGreen());

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2010);
	document->Close();
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Footnote.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetPositionAndNumberFormat.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(0);

	//Set the number format, restart rule and position for the footnote
	sec->GetFootnoteOptions()->SetNumberFormat(FootnoteNumberFormat::UpperCaseLetter);
	sec->GetFootnoteOptions()->SetRestartRule(FootnoteRestartRule::RestartPage);
	sec->GetFootnoteOptions()->SetPosition(FootnotePosition::PrintAsEndOfSection);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}

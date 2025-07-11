#include "pch.h"
using namespace Spire::Doc;


int main() {
	

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BlankTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"EmbedPrivateFont.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section and add a paragraph
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Paragraph> p = section->AddParagraph();

	//Append text to the paragraph, then set the font name and font size
	intrusive_ptr<TextRange> range = p->AppendText(L"Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.");
	range->GetCharacterFormat()->SetFontName(L"PT Serif Caption");
	range->GetCharacterFormat()->SetFontSize(20);

	//Allow embedding font in document
	doc->SetEmbedFontsInFile(true);

	//Embed private font from font file into the document
	intrusive_ptr<PrivateFontPath> tempVar = new PrivateFontPath(L"PT Serif Caption", (input_path + L"PT_Serif-Caption-Web-Regular.ttf").c_str());
	doc->GetPrivateFontList().push_back(tempVar);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
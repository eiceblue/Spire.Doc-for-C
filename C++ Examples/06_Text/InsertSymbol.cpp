#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertSymbol.docx";

	//Create Word document.
	intrusive_ptr<Document> document =  new Document();

	//Add a section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a paragraph.
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Use unicode characters to create symbol Ä.
	std::wstring tempA = L"\u00c4";
	intrusive_ptr<TextRange> tr = paragraph->AppendText(tempA.c_str());

	//Set the color of symbol Ä.
	tr->GetCharacterFormat()->SetTextColor(Color::GetRed());

	//Add symbol Ë.
	std::wstring tempB = L"\u00cb";
	paragraph->AppendText(tempB.c_str());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}

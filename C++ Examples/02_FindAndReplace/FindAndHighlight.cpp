#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+ L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path+L"FindAndHighlight.docx";
	
	//Create word document
	intrusive_ptr<Document> document = new Document();
	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());
	//Find text
	std::vector<intrusive_ptr<TextSelection>> textSelections = document->FindAllString(L"word", false, true);

	//Set hightlight
	for (intrusive_ptr<TextSelection> selection : textSelections)
	{
		selection->GetAsOneRange()->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());
	}

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();


}


#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"ApplyEmphasisMark.docx";

	// Create a new document and load from file;
	intrusive_ptr<Document> doc =  new Document();
	doc->LoadFromFile(inputFile.c_str());
	//Find text to emphasize
	std::vector<intrusive_ptr<TextSelection>> textSelections = doc->FindAllString(L"Spire.Doc for .NET", false, true);

	//Set emphasis mark to the found text
	for (intrusive_ptr<TextSelection> selection : textSelections)
	{
		selection->GetAsOneRange()->GetCharacterFormat()->SetEmphasisMark(Emphasis::Dot);
	}

	//Save and launch the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
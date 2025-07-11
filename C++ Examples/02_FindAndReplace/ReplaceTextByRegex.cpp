#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReplaceTextByRegex.docx";
	wstring outputFile = output_path + L"ReplaceTextByRegex.docx";
	//create a document
	intrusive_ptr<Document> doc = new Document();

	//Load the document from disk.
	doc->LoadFromFile(inputFile.c_str());

	//create a regex, match the text that starts with #
	intrusive_ptr<Regex> regex = new Regex(L"\\#\\w+\\b");

	//replace the text by regex
	doc->Replace(regex, L"Spire.Doc");

	//save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}
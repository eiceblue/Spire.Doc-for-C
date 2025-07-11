#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"InputHtml.txt";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlStringToWord.docx";

	//Get html string.
	ifstream in(inputFile.c_str(), ios::in);
	istreambuf_iterator<char> begin(in), end;
	wstring HTML(begin, end);
	in.close();

	//Create a new document.
	intrusive_ptr<Document> document = new Document();

	//Add a section.
	intrusive_ptr<Section> sec = document->AddSection();

	//Add a paragraph and append html string.
	intrusive_ptr<Paragraph> para = sec->AddParagraph();
	para->AppendHTML(HTML.c_str());

	//Save it to a Word file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Toc.docx";
	wstring inputFile_2 = input_path + L"Template_N3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyDocumentStyles.docx";

	//Load source document from disk
	intrusive_ptr<Document> srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile.c_str());


	//Load destination document from disk
	intrusive_ptr<Document> destDoc = new Document();
	destDoc->LoadFromFile(inputFile_2.c_str());

	//Get the style collections of source document
	intrusive_ptr<StyleCollection> styles = srcDoc->GetStyles();

	//Add the style to destination document
	for (int i = 0; i < styles->GetCount(); i++)
	{
		intrusive_ptr<IStyle> style = styles->GetItem(i);
		intrusive_ptr<Style> destStyle = Object::Dynamic_cast<Style>(style);
		destDoc->GetStyles()->Add(destStyle);
	}

	//Save the Word file
	destDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	destDoc->Close();

}

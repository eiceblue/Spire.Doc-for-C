#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile_1 = input_path + L"SupportDocumentCompare1.docx";
	wstring inputFile_2 = input_path + L"SupportDocumentCompare2.docx";
	wstring outputFile = output_path + L"CompareDocuments.docx";

	//Load the first document
	intrusive_ptr<Document> doc1 =  new Document();
	doc1->LoadFromFile(inputFile_1.c_str());
	//Load the second document
	intrusive_ptr<Document> doc2 =  new Document();
	doc2->LoadFromFile(inputFile_2.c_str());
	//Compare the two documents
	doc1->Compare(doc2, L"E-iceblue");

	//Save as docx file.
	doc1->SaveToFile(outputFile.c_str(), Spire::Doc::FileFormat::Docx2013);
	doc1->Close();
	doc2->Close();

}
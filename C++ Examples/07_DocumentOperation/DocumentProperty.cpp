#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Summary_of_Science.doc";
	wstring outputFile = output_path + L"DocumentProperty.docx";

	//Open a blank Word document as template.
	intrusive_ptr<Document> document =  new Document();
	document->LoadFromFile(inputFile.c_str());

	document->GetBuiltinDocumentProperties()->SetTitle(L"Document Demo Document");
	document->GetBuiltinDocumentProperties()->SetSubject(L"demo");
	document->GetBuiltinDocumentProperties()->SetAuthor(L"James");
	document->GetBuiltinDocumentProperties()->SetCompany(L"e-iceblue");
	document->GetBuiltinDocumentProperties()->SetManager(L"Jakson");
	document->GetBuiltinDocumentProperties()->SetCategory(L"Doc Demos");
	document->GetBuiltinDocumentProperties()->SetKeywords(L"Document, Property, Demo");
	document->GetBuiltinDocumentProperties()->SetComments(L"This document is just a demo.");

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
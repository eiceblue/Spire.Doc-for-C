#include "pch.h"


using namespace Spire::Doc;

int main() 
{
	std::wstring outputFile = OUTPUTPATH"HideEmptyRegions.docx";
	std::wstring inputFile = DATAPATH"HideEmptyRegions.doc";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	std::vector<LPCWSTR_S> filedNames = { L"Contact Name", L"Fax", L"Date" };
	//C# TO C++ CONVERTER TODO TASK: There is no C++ equivalent to 'ToString':
	std::vector<LPCWSTR_S> filedValues = { L"John Smith", L"+1 (69) 123456", DateTime::GetNow()->GetDate()->ToString() };
	//Set the value to remove paragraphs which contain empty field.
	document->GetMailMerge()->SetHideEmptyParagraphs(true);
	//Set the value to remove group which contain empty field.
	document->GetMailMerge()->SetHideEmptyGroup(true);
	document->GetMailMerge()->Execute(filedNames, filedValues);
	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
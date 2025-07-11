#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddVariables.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a paragraph.
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Add a DocVariable field.
	paragraph->AppendField(L"A1", FieldType::FieldDocVariable);

	//Add a document variable to the DocVariable field.
	document->GetVariables()->Add(L"A1", L"12");

	//Update fields.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
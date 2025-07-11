#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveField.docx";
	std::wstring inputFile = DATAPATH"/IfFieldSample.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first field
	intrusive_ptr<Field> field = document->GetFields()->GetItem(0);

	//Get the paragraph of the field
	intrusive_ptr<Paragraph> par = field->GetOwnerParagraph();
	//Get the index of the  field
	int index = par->GetChildObjects()->IndexOf(field);
	//Remove if field via index
	par->GetChildObjects()->RemoveAt(index);

	//Save doc file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
#include "pch.h"
#include <regex>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ReplaceTextInTable.docx";
	std::wstring inputFile = DATAPATH"/ReplaceTextInTable.docx";

	//Load Word from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Define a regular expression to match the {} with its content
	intrusive_ptr<Regex> regex = new Regex(L"{[^\\}]+\\}");

	//Replace the text of table with regex
	table->Replace(regex, L"E-iceblue");

	//Replace old text with new text in table
	table->Replace(L"Beijing", L"Component", false, true);

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}

#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring inputFile = DATAPATH"/RemoveReadOnlyRestriction.docx";
	std::wstring outputFile = OUTPUTPATH"/RemoveReadOnlyRestriction.docx";

	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	//Remove ReadOnly Restriction.
	doc->Protect(ProtectionType::NoProtection);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}

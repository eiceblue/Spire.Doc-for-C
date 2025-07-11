#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetSuperscriptAndSubscript.docx";
	//Create word document
	intrusive_ptr<Document> document =  new Document();

	//Create a new section
	intrusive_ptr<Section> section = document->AddSection();

	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->AppendText(L"E = mc");
	intrusive_ptr<TextRange> range1 = paragraph->AppendText(L"2");

	//Set supperscript
	range1->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SuperScript);

	paragraph->AppendBreak(BreakType::LineBreak);
	paragraph->AppendText(L"F");
	intrusive_ptr<TextRange> range2 = paragraph->AppendText(L"n");

	//Set subscript
	range2->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SubScript);

	paragraph->AppendText(L" = F");
	paragraph->AppendText(L"n-1")->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SubScript);
	paragraph->AppendText(L" + F");
	paragraph->AppendText(L"n-2")->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SubScript);

	//Set font size
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> item = paragraph->GetChildObjects()->GetItem(i);
		if (Object::CheckType<TextRange>(item))
		{
			intrusive_ptr<TextRange> tr = boost::dynamic_pointer_cast<TextRange>(item);
			tr->GetCharacterFormat()->SetFontSize(36);
		}
	}

	//Save and launch the document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
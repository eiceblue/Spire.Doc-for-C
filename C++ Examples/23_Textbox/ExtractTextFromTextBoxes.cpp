#include "pch.h"
#include <locale>
#include <codecvt>
using namespace Spire::Doc;

void ExtractTextFromTables(intrusive_ptr<Table> table, wofstream& sw)
{
	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
		for (int j = 0; j < row->GetCells()->GetCount(); j++)
		{
			intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(j);
			for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
			{
				intrusive_ptr<Paragraph> paragraph = cell->GetParagraphs()->GetItemInParagraphCollection(k);
				sw << (paragraph->GetText());
			}
		}
	}
}
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractTextFromTextBoxes.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractTextFromTextBoxes.txt";

	//Create a Word document.
	intrusive_ptr<Document>  document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Verify whether the document contains a textbox or not.
	if (document->GetTextBoxes()->GetCount() > 0)
	{
		wofstream sw(outputFile);
		auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
		sw.imbue(LocUtf8);
		//Traverse the document.
		for (int i = 0; i < document->GetSections()->GetCount(); i++)
		{
			intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
			for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
			{
				intrusive_ptr<Paragraph> p = section->GetParagraphs()->GetItemInParagraphCollection(j);
				for (int k = 0; k < p->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> obj = p->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::TextBox)
					{
						intrusive_ptr<TextBox> textbox = Object::Dynamic_cast<TextBox>(obj);
						for (int l = 0; l < textbox->GetChildObjects()->GetCount(); l++)
						{
							intrusive_ptr<DocumentObject> objt = textbox->GetChildObjects()->GetItem(l);
							//Extract text from paragraph in TextBox.
							if (objt->GetDocumentObjectType() == DocumentObjectType::Paragraph)
							{
								std::wstring tempStr = (Object::Dynamic_cast<Paragraph>(objt))->GetText();
								sw << tempStr;
							}

							//Extract text from Table in TextBox.
							if (objt->GetDocumentObjectType() == DocumentObjectType::Table)
							{
								intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(objt);
								ExtractTextFromTables(table, sw);
							}
						}
					}
				}
			}

		}
		sw.close();
	}
	document->Close();
}
    
	//delete document;
	



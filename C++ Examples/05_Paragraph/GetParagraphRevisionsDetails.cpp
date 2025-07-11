#include "pch.h"


using namespace Spire::Doc;

wstring GetRevisionType(EditRevisionType type)
{
	switch (type)
	{
	case EditRevisionType::Deletion:
		return L"Deletion";
		break;
	case EditRevisionType::Insertion:
		return L"Insertion";
		break;
	}
	return L"";
}

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Revisions.docx";
	wstring outputFile = output_path + L"GetParagraphRevisionsDetails.txt";
	
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	wstring* builder = new wstring();

	//loop paragraph
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			int sectionIndex = document->GetIndex(section);
			int paragraphIndex = section->GetIndex(paragraph);
			if (paragraph->GetIsDeleteRevision())
			{

				builder->append(L"The section " + to_wstring(sectionIndex) + L" paragraph " + to_wstring(paragraphIndex) + L" has been changed (deleted).\n");
				std::wstring author = paragraph->GetDeleteRevision()->GetAuthor();
				builder->append(L"Author: " + author + L"\n");
				intrusive_ptr<DateTime> time = paragraph->GetDeleteRevision()->GetDateTime();
				builder->append(L"DateTime: ");
				builder->append(time->ToString());
				//builder.append(time->ToLongTimeString());
				builder->append(L"\n");
				std::wstring type = GetRevisionType(paragraph->GetDeleteRevision()->GetType());
				builder->append(L"Type: " + type + L"\n");
				builder->append(L"");
				builder->append(L"\n");
			}
			else if (paragraph->GetIsInsertRevision())
			{
				builder->append(L"The section " + to_wstring(sectionIndex) + L" paragraph " + to_wstring(paragraphIndex) + L" has been changed (inserted).\n");
				std::wstring author = paragraph->GetInsertRevision()->GetAuthor();
				builder->append(L"Author: " + author + L"\n");
				intrusive_ptr<DateTime> time = paragraph->GetInsertRevision()->GetDateTime();
				builder->append(L"DateTime: ");
				std::wstring timeString = time->ToString();
				//wstring timeString = time->ToLongDateString();
				builder->append(timeString.c_str());
				builder->append(L"\n");
				std::wstring type = GetRevisionType(paragraph->GetInsertRevision()->GetType());
				builder->append(L"Type: " + type + L"\n");
				builder->append(L"");
				builder->append(L"\n");
			}
			else
			{
				for (int k = 0; k < paragraph->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(k);
					if (obj->GetDocumentObjectType() == DocumentObjectType::TextRange)
					{
						intrusive_ptr<TextRange> textRange = Object::Dynamic_cast<TextRange>(obj);

						if (textRange->GetIsDeleteRevision())
						{
							builder->append(L"The section " + to_wstring(sectionIndex) + L" paragraph " + to_wstring(paragraphIndex) + L" textrange " + to_wstring(paragraph->GetIndex(textRange)) + L" has been changed (deleted).\n");
							std::wstring author = textRange->GetDeleteRevision()->GetAuthor();
							builder->append(L"Author: " + author + L"\n");
							intrusive_ptr<DateTime> time = textRange->GetDeleteRevision()->GetDateTime();
							builder->append(L"DateTime: ");
							builder->append(time->ToString());
							builder->append(L"\n");
							std::wstring type = GetRevisionType(textRange->GetDeleteRevision()->GetType());
							builder->append(L"Type: " + type + L"\n");
							builder->append(L"Change Text: ");
							builder->append(textRange->GetText());
							builder->append(L"");
							builder->append(L"\n");
						}
						else if (textRange->GetIsInsertRevision())
						{
							builder->append(L"The section " + to_wstring(sectionIndex) + L" paragraph " + to_wstring(paragraphIndex) + L" textrange " + to_wstring(paragraph->GetIndex(textRange)) + L" has been changed (inserted).\n");
							std::wstring author = textRange->GetInsertRevision()->GetAuthor();
							builder->append(L"Author: " + author + L"\n");
							intrusive_ptr<DateTime> time = textRange->GetInsertRevision()->GetDateTime();
							builder->append(L"DateTime: ");
							builder->append(time->ToString());
							builder->append(L"\n");
							std::wstring type = GetRevisionType(textRange->GetInsertRevision()->GetType());
							builder->append(L"Type: " + type + L"\n");
							builder->append(L"Change Text: ");
							builder->append(textRange->GetText());
							builder->append(L"");
							builder->append(L"\n");
						}

					}
				}
			}
		}
	}

	//Save to file.
	wofstream write(outputFile.c_str());
	write << builder->c_str();
	write.close();
	document->Close();


}
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SplitWordFileByPageBreak.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SplitDocByPageBreak/";

	//Create Word document.
	intrusive_ptr<Document> original = new Document();

	//Load the file from disk.
	original->LoadFromFile(inputFile.c_str());

	//Create a new word document and add a section to it.
	intrusive_ptr<Document> newWord = new Document();
	intrusive_ptr<Section> section = newWord->AddSection();
	original->CloneDefaultStyleTo(newWord);
	original->CloneThemesTo(newWord);
	original->CloneCompatibilityTo(newWord);

	//Split the original word document into separate documents according to page break.
	int index = 0;

	//Traverse through all sections of original document.
	int sectionCount = original->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		//Traverse through all GetBody() child objects of each section.
		intrusive_ptr<Section> sec = original->GetSections()->GetItemInSectionCollection(i);

		int ChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < ChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> obj = sec->GetBody()->GetChildObjects()->GetItem(j);
			if (Object::CheckType<Paragraph>(obj))
			{
				intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(obj);
				sec->CloneSectionPropertiesTo(section);
				//Add paragraph object in original section into section of new document.
				section->GetBody()->GetChildObjects()->Add(para->Clone());

				int parObjCount = para->GetChildObjects()->GetCount();
				for (int k = 0; k < parObjCount; k++)
				{
					intrusive_ptr<DocumentObject> parobj = para->GetChildObjects()->GetItem(k);
					if (Object::CheckType<Break>(parobj) && (Object::Dynamic_cast<Break>(parobj))->GetBreakType() == BreakType::PageBreak)
					{
						//Get the index of page break in paragraph.
						int i = para->GetChildObjects()->IndexOf(parobj);

						//Remove the page break from its paragraph.
						section->GetBody()->GetLastParagraph()->GetChildObjects()->RemoveAt(i);

						//Save the new document to a Docx file.
						std::wstring file = outputFile + L"SplitDocByPageBreak-" + to_wstring(index) + L".docx";

						newWord->SaveToFile(file.c_str(), FileFormat::Docx);
						index++;

						//Create a new document and add a section.
						newWord = new Document();
						section = newWord->AddSection();
						original->CloneDefaultStyleTo(newWord);
						original->CloneThemesTo(newWord);
						original->CloneCompatibilityTo(newWord);
						sec->CloneSectionPropertiesTo(section);
						//Add paragraph object in original section into section of new document.
						section->GetBody()->GetChildObjects()->Add(para->Clone());
						if (section->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->GetCount() == 0)
						{
							//Remove the first blank paragraph.
							section->GetBody()->GetChildObjects()->RemoveAt(0);
						}
						else
						{
							//Remove the child objects before the page break.
							while (i >= 0)
							{
								section->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->RemoveAt(i);
								i--;
							}
						}
					}
				}
			}
			if (Object::CheckType<Table>(obj))
			{
				//Add table object in original section into section of new document.
				section->GetBody()->GetChildObjects()->Add(obj->Clone());
			}
		}
	}

	//Save to file.
	std::wstring result = outputFile + L"SplitDocByPageBreak-" + to_wstring(index) + L".docx";
	newWord->SaveToFile(result.c_str(), FileFormat::Docx2013);
	newWord->Close();
}

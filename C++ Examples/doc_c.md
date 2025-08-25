# C++ Precompiled Header
## Defines library linking for Spire.Doc
```cpp
#pragma comment(lib,"../lib/Spire.Doc.Cpp.lib")
```

---

# spire.doc cpp helloworld
## create a simple word document with "Hello World!" text
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Create a new section
intrusive_ptr<Section> section = document->AddSection();

//Create a new paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Append Text
paragraph->AppendText(L"Hello World!");
```

---

# Spire.Doc CPP Find and Highlight
## Find and highlight specific text in a Word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Find text
std::vector<intrusive_ptr<TextSelection>> textSelections = document->FindAllString(L"word", false, true);

//Set highlight
for (intrusive_ptr<TextSelection> selection : textSelections)
{
	selection->GetAsOneRange()->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());
}
```

---

# spire.doc cpp document manipulation
## replace content with another document
```cpp
//Create the first document
intrusive_ptr<Document> document1 = new Document();

//Create the second document
intrusive_ptr<Document> document2 = new Document();

//Get the first section of the first document 
intrusive_ptr<Section> section1 = document1->GetSections()->GetItemInSectionCollection(0);

//Create a regex
intrusive_ptr<Regex> regex = new Regex(L"\\[MY_DOCUMENT\\]", RegexOptions::None);

//Find the text by regex
std::vector<intrusive_ptr<TextSelection>> textSections = document1->FindAllPattern(regex);

//Travel the found strings
for (intrusive_ptr<TextSelection> seletion : textSections)
{

    //Get the paragraph
    intrusive_ptr<Paragraph> para = seletion->GetAsOneRange()->GetOwnerParagraph();

    //Get textRange
    intrusive_ptr<TextRange> textRange = seletion->GetAsOneRange();

    //Get the paragraph index
    int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

    //Insert the paragraphs of document2
    for (int i = 0; i < document2->GetSections()->GetCount(); i++)
    {
        intrusive_ptr<Section> section2 = document2->GetSections()->GetItemInSectionCollection(i);
        for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
        {
            intrusive_ptr<Paragraph> paragraph = section2->GetParagraphs()->GetItemInParagraphCollection(j);
            section1->GetBody()->GetChildObjects()->Insert(index, Object::Dynamic_cast<Paragraph>(paragraph->Clone()));
        }
    }
    //Remove the found textRange
    para->GetChildObjects()->Remove(textRange);
}
```

---

# spire.doc cpp regex replace
## Replace text using regular expression in a Word document
```cpp
//create a document
intrusive_ptr<Document> doc = new Document();

//create a regex, match the text that starts with #
intrusive_ptr<Regex> regex = new Regex(L"\\#\\w+\\b");

//replace the text by regex
doc->Replace(regex, L"Spire.Doc");
```

---

# spire.doc cpp find and replace
## replace text with field
```cpp
//Find the target text
intrusive_ptr<TextSelection> selection = document->FindString(L"summary", false, true);
//Get text range
intrusive_ptr<TextRange> textRange = selection->GetAsOneRange();
//Get its owner paragraph
intrusive_ptr<Paragraph> ownParagraph = textRange->GetOwnerParagraph();
//Get the index of this text range
int rangeIndex = ownParagraph->GetChildObjects()->IndexOf(textRange);
//Remove the text range
ownParagraph->GetChildObjects()->RemoveAt(rangeIndex);
//Remove the objects which are behind the text range
std::vector<intrusive_ptr<DocumentObject>> tempList;
for (int i = rangeIndex; i < ownParagraph->GetChildObjects()->GetCount(); i++)
{
	//Add a copy of these objects into a temp list
	tempList.push_back(ownParagraph->GetChildObjects()->GetItem(rangeIndex)->Clone());
	ownParagraph->GetChildObjects()->RemoveAt(rangeIndex);
}
//Append field to the paragraph
ownParagraph->AppendField(L"MyFieldName", FieldType::FieldMergeField);
//Put these objects back into the paragraph one by one
for (intrusive_ptr<DocumentObject> obj : tempList)
{
	ownParagraph->GetChildObjects()->Add(obj);
}
```

---

# spire.doc cpp find and replace
## replace text with table in word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Return TextSection by finding the key text string "Christmas Day, December 25".
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<TextSelection> selection = document->FindString(L"Christmas Day, December 25", true, true);

//Return TextRange from TextSection, then get OwnerParagraph through TextRange.
intrusive_ptr<TextRange> range = selection->GetAsOneRange();
intrusive_ptr<Paragraph> paragraph = range->GetOwnerParagraph();

//Return the zero-based index of the specified paragraph.
intrusive_ptr<Body> body = paragraph->GetOwnerTextBody();
int index = body->GetChildObjects()->IndexOf(paragraph);

//Create a new table.
intrusive_ptr<Table> table = section->AddTable(true);
table->ResetCells(3, 3);

//Remove the paragraph and insert table into the collection at the specified index.
body->GetChildObjects()->Remove(paragraph);
body->GetChildObjects()->Insert(index, table);
```

---

# spire.doc cpp replace
## replace text with another document
```cpp
//Load a template document 
intrusive_ptr<Document> doc = new Document(inputFile_1.c_str());

//Load another document to replace text
intrusive_ptr<IDocument> replaceDoc = new Document(inputFile_2.c_str());
//Replace specified text with the other document
doc->Replace(L"Document1", replaceDoc, false, true);
```

---

# spire.doc cpp find and replace
## replace text with html content in word document
```cpp
void ReplaceWithHTML(intrusive_ptr<TextRange> textRange, std::vector<intrusive_ptr<DocumentObject>>& replacement)
{
	//textRange index
	int index = textRange->GetOwner()->GetChildObjects()->IndexOf(textRange);

	//get owener paragraph
	intrusive_ptr<Paragraph> paragraph = textRange->GetOwnerParagraph();

	//get owner text Body
	intrusive_ptr<Body> sectionBody = paragraph->GetOwnerTextBody();

	//get the index of paragraph in section
	int paragraphIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);

	int replacementIndex = -1;
	if (index == 0)
	{
		//remove the first child object
		paragraph->GetChildObjects()->RemoveAt(0);

		replacementIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);
	}
	else if (index == paragraph->GetChildObjects()->GetCount() - 1)
	{
		paragraph->GetChildObjects()->RemoveAt(index);
		replacementIndex = paragraphIndex + 1;
	}
	else
	{
		//split owner paragraph
		intrusive_ptr<Paragraph> paragraph1 = Object::Dynamic_cast<Paragraph>(paragraph->Clone());
		while (paragraph->GetChildObjects()->GetCount() > index)
		{
			paragraph->GetChildObjects()->RemoveAt(index);
		}
		int i = 0;
		int count = index + 1;
		while (i < count)
		{
			paragraph1->GetChildObjects()->RemoveAt(0);
			i += 1;
		}
		sectionBody->GetChildObjects()->Insert(paragraphIndex + 1, paragraph1);

		replacementIndex = paragraphIndex + 1;
	}

	//insert replacement
	int finalCount = replacement.size() - 1;
	for (int i = 0; i <= finalCount; i++)
	{
		sectionBody->GetChildObjects()->Insert(replacementIndex + i, replacement[i]->Clone());
	}
}

//Load the document from disk.  
intrusive_ptr<Document> document = new Document();

//collect the objects which is used to replace text
std::vector<intrusive_ptr<DocumentObject>> replacement;

//create a temporary section
intrusive_ptr<Section> tempSection = document->AddSection();

//add a paragraph to append html
intrusive_ptr<Paragraph> par = tempSection->AddParagraph();
par->AppendHTML(HTML.c_str());

//get the objects in temporary section
for (int i = 0; i < tempSection->GetBody()->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> obj = tempSection->GetBody()->GetChildObjects()->GetItem(i);
	intrusive_ptr<DocumentObject> docObj = Object::Dynamic_cast<DocumentObject>(obj);
	replacement.push_back(docObj);
}

//Find all text which will be replaced.
std::vector<intrusive_ptr<TextSelection>> selections = document->FindAllString(L"[#placeholder]", false, true);

std::vector<intrusive_ptr<TextRange>> locations;
for (intrusive_ptr<TextSelection> selection : selections)
{
	
	intrusive_ptr<TextRange> tempVar = selection->GetAsOneRange();
	locations.push_back(tempVar);
}
std::sort(locations.begin(), locations.end());

for (intrusive_ptr<TextRange> location : locations)
{
	//replace the text with HTML.c_str()
	ReplaceWithHTML(location, replacement);
}

//remove the temp section
document->GetSections()->Remove(tempSection);
```

---

# spire.doc cpp find and replace
## replace text with image in document
```cpp
//Find the string in the document
std::vector<intrusive_ptr<TextSelection>> selections = doc->FindAllString(L"E-iceblue", true, true);
int index = 0;
intrusive_ptr<TextRange> range = nullptr;

//Remove the text and replace it with Image
for (intrusive_ptr<TextSelection> selection : selections)
{
    intrusive_ptr<DocPicture> pic = new DocPicture(doc);
    pic->LoadImageSpire(imagePath);

    range = selection->GetAsOneRange();
    index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
    range->GetOwnerParagraph()->GetChildObjects()->Insert(index, pic);
    range->GetOwnerParagraph()->GetChildObjects()->Remove(range);
}
```

---

# spire.doc cpp find and replace
## replace text in word document
```cpp
using namespace Spire::Doc;

intrusive_ptr<Document> document = new Document();

//Replace text
document->Replace(L"word", L"ReplacedText", false, true);
```

---

# spire.doc cpp extract content
## extract content between paragraphs
```cpp
void ExtractBetweenParagraphs(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, int startPara, int endPara)
{
	//Extract the content
	for (int i = startPara - 1; i < endPara; i++)
	{
		//Clone the ChildObjects of source document
		intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(doobj);
	}
}
```

---

# spire.doc cpp content extraction
## extract content between paragraph styles
```cpp
void ExtractBetweenParagraphStyles(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, const std::wstring& stylename1, const std::wstring& stylename2)
{
	int startindex = 0;
	int endindex = 0;
	//travel the sections of source document

	for (int i = 0; i < sourceDocument->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = sourceDocument->GetSections()->GetItemInSectionCollection(i);
		//travel the paragraphs
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			//Judge paragraph style1
			if (paragraph->GetStyleName() == stylename1)
			{
				//Get the paragraph index
				startindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
			//Judge paragraph style2
			if (paragraph->GetStyleName() == stylename2)
			{
				//Get the paragraph index
				endindex = section->GetBody()->GetParagraphs()->IndexOf(paragraph);
			}
		}
		//Extract the content
		for (int i = startindex + 1; i < endindex; i++)
		{
			//Clone the ChildObjects of source document
			intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

			//Add to destination document 
			destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(doobj);
		}
	}
}
```

---

# spire.doc cpp extract content
## extract paragraphs based on style
```cpp
//Create a new document
intrusive_ptr<Document> document = new Document();
wstring styleName1 = L"Heading1";
wstring style1Text;

//Load file from disk
document->LoadFromFile(inputFile.c_str());

//Extract paragraph based on style
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
        if (paragraph->GetStyleName() != nullptr && paragraph->GetStyleName() == styleName1)
        {
            style1Text.append(paragraph->GetText());
        }
    }
}

document->Close();
```

---

# spire.doc cpp bookmark extraction
## extract content from bookmark in document
```cpp
//Locate the bookmark in source document
intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(sourcedocument);

//Find bookmark by name
navigator->MoveToBookmark(L"Test", true, true);

//get text Body part
intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

//Create a TextRange type list
std::vector<intrusive_ptr<TextRange>> list;

//Traverse the items of text Body
for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
	//if it is paragraph
	if (Object::CheckType<Paragraph>(item))
	{
		//Traverse the ChildObjects of the paragraph
		for (int i = 0; i < (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetCount(); i++)
		{
			intrusive_ptr<DocumentObject> childObject = (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetItem(i);
			//if it is TextRange
			if (Object::CheckType<TextRange>(childObject))
			{
				//Add it into list
				intrusive_ptr<TextRange> range = boost::dynamic_pointer_cast<TextRange>(childObject);
				list.push_back(range);
			}
		}
	}
}

//Add the extract content to destinationDoc document
for (size_t m = 0; m < list.size(); m++)
{
	paragraph->GetChildObjects()->Add(list[m]->Clone());
}
```

---

# spire.doc cpp comment extraction
## extract content from comment range in document
```cpp
//Create a document
intrusive_ptr<Document> sourceDoc = new Document();

//Create a destination document
intrusive_ptr<Document> destinationDoc = new Document();

//Add section for destination document
intrusive_ptr<Section> destinationSec = destinationDoc->AddSection();

//Get the first comment
intrusive_ptr<Comment> comment = sourceDoc->GetComments()->GetItem(0);

//Get the paragraph of obtained comment
intrusive_ptr<Paragraph> para = comment->GetOwnerParagraph();

//Get index of the CommentMarkStart 
int startIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkStart());

//Get index of the CommentMarkEnd
int endIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkEnd());

//Traverse paragraph ChildObjects
for (int i = startIndex; i <= endIndex; i++)
{
    //Clone the ChildObjects of source document
    intrusive_ptr<DocumentObject> doobj = para->GetChildObjects()->GetItem(i)->Clone();

    //Add to destination document 
    destinationSec->AddParagraph()->GetChildObjects()->Add(doobj);
}
```

---

# spire.doc cpp extract content
## extract content from paragraph to table
```cpp
void ExtractByTable(intrusive_ptr<Document> sourceDocument, intrusive_ptr<Document> destinationDocument, int startPara, int tableNo)
{
	//Get the table from the source document
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(tableNo - 1));

	//Get the table index
	int index = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(table);
	for (int i = startPara - 1; i <= index; i++)
	{
		//Clone the ChildObjects of source document
		intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(doobj);
	}
}
```

---

# spire.doc cpp extract content
## extract document content starting from form field
```cpp
//Create the source document
intrusive_ptr<Document> sourceDocument = new Document();

//Create a destination document
intrusive_ptr<Document> destinationDoc = new Document();

//Add a section
intrusive_ptr<Section> section = destinationDoc->AddSection();

//Define a variables
int index = 0;

//Traverse FormFields
for (int i = 0; i < sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount(); i++)
{
	intrusive_ptr<FormField> field = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);
	//Find FieldFormTextInput type field
	if (field->GetType() == FieldType::FieldFormTextInput)
	{
		//Get the paragraph
		intrusive_ptr<Paragraph> paragraph = field->GetOwnerParagraph();

		//Get the index
		index = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(paragraph);
		break;
	}
}

//Extract the content
for (int i = index; i < index + 3; i++)
{
	//Clone the ChildObjects of source document
	intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

	//Add to destination document 
	section->GetBody()->GetChildObjects()->Add(doobj);
}
```

---

# spire.doc cpp sections
## add and delete sections in a word document
```cpp
//Add a section
doc->AddSection();
//Delete the last section
doc->GetSections()->RemoveAt(doc->GetSections()->GetCount() - 1);
```

---

# spire.doc cpp section cloning
## clone sections from one document to another
```cpp
//Create source and destination documents
intrusive_ptr<Document> srcDoc = new Document();
intrusive_ptr<Document> desDoc = new Document();

intrusive_ptr<Section> cloneSection = nullptr;
for (int i = 0; i < srcDoc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = srcDoc->GetSections()->GetItemInSectionCollection(i);
    //Clone section
    cloneSection = section->CloneSection();
    //Add the cloneSection in destination file
    desDoc->GetSections()->Add(cloneSection);
}
```

---

# spire.doc cpp section
## clone section content
```cpp
//Get the first section
intrusive_ptr<Section> sec1 = doc->GetSections()->GetItemInSectionCollection(0);
//Get the second section
intrusive_ptr<Section> sec2 = doc->GetSections()->GetItemInSectionCollection(1);

//Loop through the contents of sec1
for (int i = 0; i < sec1->GetBody()->GetChildObjects()->GetCount(); i++)
{
    intrusive_ptr<DocumentObject> obj = sec1->GetBody()->GetChildObjects()->GetItem(i);
    //Clone the contents to sec2
    sec2->GetBody()->GetChildObjects()->Add(obj->Clone());
}
```

---

# spire.doc cpp page setup
## modify page setup of sections
```cpp
//Loop through all sections
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    //Modify the margins
    section->GetPageSetup()->SetMargins(new MarginsF(100, 80, 100, 80));
    //Modify the page size
    section->GetPageSetup()->SetPageSize(PageSize::Letter());
}
```

---

# spire.doc cpp section
## remove content from document sections
```cpp
//Loop through all sections
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    //Remove header content
    section->GetHeadersFooters()->GetHeader()->GetChildObjects()->Clear();
    //Remove GetBody() content
    section->GetBody()->GetChildObjects()->Clear();
    //Remove footer content
    section->GetHeadersFooters()->GetFooter()->GetChildObjects()->Clear();
}
```

---

# spire.doc cpp paragraph tab stops
## add tab stops to paragraphs in Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Add a section.
intrusive_ptr<Section> section = document->AddSection();

//Add paragraph 1.
intrusive_ptr<Paragraph> paragraph1 = section->AddParagraph();

//Add tab and set its position (in points).
intrusive_ptr<Tab> tab = paragraph1->GetFormat()->GetTabs()->AddTab(28);

//Set tab alignment.
tab->SetJustification(TabJustification::Left);

//Move to next tab and append text.
paragraph1->AppendText(L"\tWashing Machine");

//Add another tab and set its position (in points).
tab = paragraph1->GetFormat()->GetTabs()->AddTab(280);

//Set tab alignment.
tab->SetJustification(TabJustification::Left);

//Specify tab leader type.
tab->SetTabLeader(TabLeader::Dotted);

//Move to next tab and append text.
paragraph1->AppendText(L"\t$650");

//Add paragraph 2.
intrusive_ptr<Paragraph> paragraph2 = section->AddParagraph();

//Add tab and set its position (in points).
tab = paragraph2->GetFormat()->GetTabs()->AddTab(28);

//Set tab alignment.
tab->SetJustification(TabJustification::Left);

//Move to next tab and append text.
paragraph2->AppendText(L"\tRefrigerator");

//Add another tab and set its position (in points).
tab = paragraph2->GetFormat()->GetTabs()->AddTab(280);

//Set tab alignment.
tab->SetJustification(TabJustification::Left);

//Specify tab leader type.
tab->SetTabLeader(TabLeader::NoLeader);

//Move to next tab and append text.
paragraph2->AppendText(L"\t$800");
```

---

# spire.doc cpp paragraph
## allow latin text wrap in middle of word
```cpp
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Paragraph> para = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);
//Allow Latin text to wrap in the middle of a word
para->GetFormat()->SetWordWrap(false);
```

---

# spire.doc cpp paragraph
## copy paragraph between documents
```cpp
using namespace Spire::Doc;

//Create Word document1.
intrusive_ptr<Document> document1 = new Document();

//Create a new document.
intrusive_ptr<Document> document2 = new Document();

//Get paragraph 1 and paragraph 2 in document1.
intrusive_ptr<Section> s = document1->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> p1 = s->GetParagraphs()->GetItemInParagraphCollection(0);
intrusive_ptr<Paragraph> p2 = s->GetParagraphs()->GetItemInParagraphCollection(1);

//Copy p1 and p2 to document2.
intrusive_ptr<Section> s2 = document2->AddSection();
intrusive_ptr<Paragraph> NewPara1 = Object::Dynamic_cast<Paragraph>(p1->Clone());
s2->GetParagraphs()->Add(NewPara1);
intrusive_ptr<Paragraph> NewPara2 = Object::Dynamic_cast<Paragraph>(p2->Clone());
s2->GetParagraphs()->Add(NewPara2);
```

---

# spire.doc c++ catalog
## create a catalog with heading styles and list formatting
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Add a new section. 
intrusive_ptr<Section> section = document->AddSection();
intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();

//Add Heading 1.
paragraph = section->AddParagraph();
paragraph->AppendText(L"Heading1");
paragraph->ApplyStyle(BuiltinStyle::Heading1);
paragraph->GetListFormat()->ApplyNumberedStyle();

//Add Heading 2.
paragraph = section->AddParagraph();
paragraph->AppendText(L"Heading2");
paragraph->ApplyStyle(BuiltinStyle::Heading2);

//List style for Headings 2.
intrusive_ptr<ListStyle> listSty2 = new ListStyle(document, ListType::Numbered);

for (int i = 0; i < listSty2->GetLevels()->GetCount(); i++)
{
	intrusive_ptr<ListLevel> listLev = listSty2->GetLevels()->GetItem(i);
	listLev->SetUsePrevLevelPattern(true);
	listLev->SetNumberPrefix(L"1.");
}
listSty2->SetName(L"MyStyle2");
document->GetListStyles()->Add(listSty2);
paragraph->GetListFormat()->ApplyStyle(listSty2->GetName());

//Add list style 3.
intrusive_ptr<ListStyle> listSty3 = new ListStyle(document, ListType::Numbered);

for (int i = 0; i < listSty3->GetLevels()->GetCount(); i++)
{
	intrusive_ptr<ListLevel> listlev = listSty3->GetLevels()->GetItem(i);
	listlev->SetUsePrevLevelPattern(true);
	listlev->SetNumberPrefix(L"1.1.");
}
listSty3->SetName(L"MyStyle3");
document->GetListStyles()->Add(listSty3);

//Add Heading 3.
for (int i = 0; i < 4; i++)
{
	paragraph = section->AddParagraph();

	//Append text
	paragraph->AppendText(L"Heading3");

	//Apply list style 3 for Heading 3
	paragraph->ApplyStyle(BuiltinStyle::Heading3);
	paragraph->GetListFormat()->ApplyStyle(listSty3->GetName());
}
```

---

# spire.doc cpp paragraph
## get paragraphs by style name
```cpp
// Iterate through all sections in the document
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
    
    // Iterate through all paragraphs in the section
    for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
        
        // Check if the paragraph has the desired style name
        wstring style_name = paragraph->GetStyleName();
        if (style_name.compare(L"Heading1") == 0)
        {
            // Process paragraph with "Heading1" style
            wstring paragraphText = paragraph->GetText();
        }
    }
}
```

---

# Spire.Doc C++ Paragraph Revisions
## Extract details of paragraph revisions in a Word document
```cpp
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

// Main revision checking logic
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
			std::wstring author = paragraph->GetDeleteRevision()->GetAuthor();
			intrusive_ptr<DateTime> time = paragraph->GetDeleteRevision()->GetDateTime();
			std::wstring type = GetRevisionType(paragraph->GetDeleteRevision()->GetType());
		}
		else if (paragraph->GetIsInsertRevision())
		{
			std::wstring author = paragraph->GetInsertRevision()->GetAuthor();
			intrusive_ptr<DateTime> time = paragraph->GetInsertRevision()->GetDateTime();
			std::wstring type = GetRevisionType(paragraph->GetInsertRevision()->GetType());
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
						std::wstring author = textRange->GetDeleteRevision()->GetAuthor();
						intrusive_ptr<DateTime> time = textRange->GetDeleteRevision()->GetDateTime();
						std::wstring type = GetRevisionType(textRange->GetDeleteRevision()->GetType());
					}
					else if (textRange->GetIsInsertRevision())
					{
						std::wstring author = textRange->GetInsertRevision()->GetAuthor();
						intrusive_ptr<DateTime> time = textRange->GetInsertRevision()->GetDateTime();
						std::wstring type = GetRevisionType(textRange->GetInsertRevision()->GetType());
					}
				}
			}
		}
	}
}
```

---

# spire.doc cpp paragraph
## hide paragraph text in document
```cpp
//Get the first section and the first paragraph from the word document.
intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(0);

//Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
for (int i = 0; i < para->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(i);
	if (Object::CheckType<TextRange>(obj))
	{
		intrusive_ptr<TextRange> range = boost::dynamic_pointer_cast<TextRange>(obj);
		range->GetCharacterFormat()->SetHidden(true);
	}
}
```

---

# spire.doc cpp rtf
## Insert RTF string into Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Add a new section.
intrusive_ptr<Section> section = document->AddSection();

//Add a paragraph to the section.
intrusive_ptr<Paragraph> para = section->AddParagraph();

//Declare a String variable to store the Rtf string.
std::wstring rtfString = L"{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}";

//Append Rtf string to paragraph.
para->AppendRTF(rtfString.c_str());
```

---

# spire.doc cpp pagination
## manage paragraph pagination
```cpp
//Get the first section and the paragraph we want to manage the pagination.
intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(4);

//Set the pagination format as Format.PageBreakBefore for the checked paragraph.
para->GetFormat()->SetPageBreakBefore(true);
```

---

# spire.doc cpp paragraph
## remove all paragraphs from document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Remove paragraphs from every section in the document
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
	section->GetParagraphs()->Clear();
}
```

---

# spire.doc cpp remove empty lines
## remove empty paragraphs from Word document
```cpp
// Traverse every section on the word document and remove the null and empty paragraphs
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> secChildObject = section->GetBody()->GetChildObjects()->GetItem(j);
        if (secChildObject->GetDocumentObjectType() == DocumentObjectType::Paragraph)
        {
            intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(secChildObject);
            std::wstring paraText = para->GetText();
            if (paraText.empty())
            {
                section->GetBody()->GetChildObjects()->Remove(secChildObject);
                j--;
            }
        }
    }
}
```

---

# spire.doc cpp paragraph
## remove specific paragraph from document
```cpp
//Remove the first paragraph from the first section of the document.
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->RemoveAt(0);
```

---

# Spire.Doc C++ Frame Position
## Set frame position in a Word document
```cpp
//Get a paragraph
intrusive_ptr<Paragraph> paragraph = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

//Set the Frame's position
if (paragraph->GetIsFrame())
{
    paragraph->GetFrame()->SetHorizontalPosition(150.0f);
    paragraph->GetFrame()->SetVerticalPosition(150.0f);
}
```

---

# spire.doc cpp paragraph shading
## set paragraph background color and text background color
```cpp
//Create Word document.
intrusive_ptr<Document> document =  new Document();

//Get a paragraph.
intrusive_ptr<Paragraph> paragaph = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

//Set background color for the paragraph.
paragaph->GetFormat()->SetBackColor(Color::GetYellow());

//Set background color for the selected text of paragraph.
paragaph = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(2);
intrusive_ptr<TextSelection> selection = paragaph->Find(L"Christmas", true, false);
intrusive_ptr<TextRange> range = selection->GetAsOneRange();
range->GetCharacterFormat()->SetTextBackgroundColor(Color::GetYellow());
```

---

# spire.doc cpp paragraph formatting
## set space between Asian and Latin text
```cpp
intrusive_ptr<Paragraph> para = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

//Set whether to automatically adjust space between Asian text and Latin text
para->GetFormat()->SetAutoSpaceDE(false);
//Set whether to automatically adjust space between Asian text and numbers
para->GetFormat()->SetAutoSpaceDN(true);
```

---

# spire.doc cpp paragraph spacing
## Set paragraph spacing in Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Create a paragraph
intrusive_ptr<Paragraph> para = new Paragraph(document);

//set the spacing before and after.
para->GetFormat()->SetBeforeAutoSpacing(false);
para->GetFormat()->SetBeforeSpacing(10);
para->GetFormat()->SetAfterAutoSpacing(false);
para->GetFormat()->SetAfterSpacing(10);

//insert the added paragraph to the word document.
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->Insert(1, para);
```

---

# spire.doc cpp emphasis mark
## apply emphasis mark to text in document
```cpp
//Find text to emphasize
std::vector<intrusive_ptr<TextSelection>> textSelections = doc->FindAllString(L"Spire.Doc for .NET", false, true);

//Set emphasis mark to the found text
for (intrusive_ptr<TextSelection> selection : textSelections)
{
    selection->GetAsOneRange()->GetCharacterFormat()->SetEmphasisMark(Emphasis::Dot);
}
```

---

# spire.doc cpp text case conversion
## Change text case in Word document to AllCaps and SmallCaps
```cpp
// Create a new document and load from file
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());
intrusive_ptr<TextRange> textRange;

//Get the first paragraph and set its CharacterFormat to AllCaps
intrusive_ptr<Paragraph> para1 = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(1);

for (int i = 0; i < para1->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> obj = para1->GetChildObjects()->GetItem(i);
	if (Object::CheckType<TextRange>(obj))
	{
		textRange = boost::dynamic_pointer_cast<TextRange>(obj);
		textRange->GetCharacterFormat()->SetAllCaps(true);
	}
}

//Get the third paragraph and set its CharacterFormat to IsSmallCaps
intrusive_ptr<Paragraph> para2 = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(3);
for (int i = 0; i < para2->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> obj = para2->GetChildObjects()->GetItem(i);
	if (Object::CheckType<TextRange>(obj))
	{
		textRange = boost::dynamic_pointer_cast<TextRange>(obj);
		textRange->GetCharacterFormat()->SetIsSmallCaps(true);
	}
}
```

---

# spire.doc cpp barcode
## create barcode in document
```cpp
//Create a document
intrusive_ptr<Document> doc =  new Document();

//Add a paragraph
intrusive_ptr<Paragraph> p = doc->AddSection()->AddParagraph();

//Add barcode and set its format
intrusive_ptr<TextRange> txtRang = p->AppendText(L"H63TWX11072");
//Set barcode font name, note you need to install the barcode font on your system at first
txtRang->GetCharacterFormat()->SetFontName(L"C39HrP60DlTt");
txtRang->GetCharacterFormat()->SetFontSize(80);
txtRang->GetCharacterFormat()->SetTextColor(Color::GetSeaGreen());
```

---

# spire.doc cpp text extraction
## extract text from document
```cpp
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());
//get text from document
wstring text = document->GetText();
```

---

# spire.doc cpp text insertion
## insert new text after found text and highlight it
```cpp
//Find all the text string "Word" from the sample document
std::vector<intrusive_ptr<TextSelection>> selections = doc->FindAllString(L"Word", true, true);
int index = 0;

//Defines text range
intrusive_ptr<TextRange> range = new TextRange(doc);

//Insert new text string after the searched text string
for (intrusive_ptr<TextSelection> selection : selections)
{
    range = selection->GetAsOneRange();
    intrusive_ptr<TextRange> newrange = new TextRange(doc);
    newrange->SetText(L"New text)");
    index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
    range->GetOwnerParagraph()->GetChildObjects()->Insert(index + 1, newrange);
}

//Find and highlight the newly added text string
std::vector<intrusive_ptr<TextSelection>> text = doc->FindAllString(L"New text", true, true);
for (intrusive_ptr<TextSelection> seletion : text)
{
    seletion->GetAsOneRange()->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());
}
```

---

# spire.doc cpp symbol insertion
## insert symbols into word document using unicode characters
```cpp
//Create Word document.
intrusive_ptr<Document> document =  new Document();

//Add a section.
intrusive_ptr<Section> section = document->AddSection();

//Add a paragraph.
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Use unicode characters to create symbol Ä.
std::wstring tempA = L"\u00c4";
intrusive_ptr<TextRange> tr = paragraph->AppendText(tempA.c_str());

//Set the color of symbol Ä.
tr->GetCharacterFormat()->SetTextColor(Color::GetRed());

//Add symbol Ë.
std::wstring tempB = L"\u00cb";
paragraph->AppendText(tempB.c_str());
```

---

# Spire.Doc C++ Text Encoding
## Load text file with specific encoding (UTF-7)
```cpp
using namespace Spire::Doc;

intrusive_ptr<Document> document =  new Document();
// Load text with UTF-7 encoding
document->LoadText(inputFile.c_str(), Encoding::GetUTF7());
document->Close();
```

---

# spire.doc cpp superscript subscript
## set superscript and subscript text in word document
```cpp
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
```

---

# spire.doc cpp text direction
## set text direction in document sections and table cells
```cpp
//Create a new document
intrusive_ptr<Document> doc =  new Document();

//Add the first section
intrusive_ptr<Section> section1 = doc->AddSection();
//Set text direction for all text in a section
section1->SetTextDirection(TextDirection::RightToLeft);

//Set text direction for a part of text
//Add the second section
intrusive_ptr<Section> section2 = doc->AddSection();
//Add a table
intrusive_ptr<Table> table = section2->AddTable();
table->ResetCells(1, 1);
intrusive_ptr<TableCell> cell = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);
//Set vertical text direction of table
cell->GetCellFormat()->SetTextDirection(TextDirection::RightToLeftRotated);
```

---

# spire.doc cpp text split
## split text into columns with line between
```cpp
//Add a column to the first section and set width and spacing
doc->GetSections()->GetItemInSectionCollection(0)->AddColumn(100.0f, 20.0f);
//Add a line between the two columns
doc->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetColumnsLineBetween(true);
```

---

# Spire.Doc C++ Language Dictionary
## Alter language dictionary in Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document =  new Document();

//Add new section and paragraph to the document.
intrusive_ptr<Section> sec = document->AddSection();
intrusive_ptr<Paragraph> para = sec->AddParagraph();

//Add a textRange for the paragraph and append some Peru Spanish words.
intrusive_ptr<TextRange> txtRange = para->AppendText(L"corrige según diccionario en inglés");
txtRange->GetCharacterFormat()->SetLocaleIdASCII(10250);

document->Close();
```

---

# Spire.Doc C++ Document Format Detection
## Detect and identify the format of a Word document
```cpp
intrusive_ptr<Document> doc = new Document();
// Load a document from file
doc->LoadFromFile(inputFile.c_str());
// Check the format info
std::wstring fileFormat = L"The file format is ";
switch (doc->GetDetectedFormatType())
{
case FileFormat::Doc:
	fileFormat += L"Microsoft Word 97-2003 document.";
	break;
case FileFormat::Dot:
	fileFormat += L"Microsoft Word 97-2003 template.";
	break;
case FileFormat::Docx:
	fileFormat += L"Office Open XML WordprocessingML Macro-Free Document.";
	break;
case FileFormat::Docm:
	fileFormat += L"Office Open XML WordprocessingML Macro-Enabled Document.";
	break;
case FileFormat::Dotx:
	fileFormat += L"Office Open XML WordprocessingML Macro-Free Template.";
	break;
case FileFormat::Dotm:
	fileFormat += L"Office Open XML WordprocessingML Macro-Enabled Template.";
	break;
case FileFormat::Rtf:
	fileFormat += L"RTF format.";
	break;
case FileFormat::WordML:
	fileFormat += L"Microsoft Word 2003 WordprocessingML format.";
	break;
case FileFormat::Html:
	fileFormat += L"HTML format.";
	break;
case FileFormat::WordXml:
	fileFormat += L"Microsoft Word xml format for word 2007-2013.";
	break;
case FileFormat::Odt:
	fileFormat += L"OpenDocument Text.";
	break;
case FileFormat::Ott:
	fileFormat += L"OpenDocument Text Template.";
	break;
case FileFormat::DocPre97:
	fileFormat += L"Microsoft Word 6 or Word 95 format.";
	break;
default:
	fileFormat += L"Unknown format.";
	break;
}
```

---

# spire.doc cpp compare
## compare two documents
```cpp
//Load the first document
intrusive_ptr<Document> doc1 = new Document();
doc1->LoadFromFile(inputFile_1.c_str());

//Load the second document
intrusive_ptr<Document> doc2 = new Document();
doc2->LoadFromFile(inputFile_2.c_str());

//Compare the two documents
doc1->Compare(doc2, L"E-iceblue");
```

---

# spire.doc cpp document comparison
## Compare two Word documents with options
```cpp
intrusive_ptr<Document> doc1 =  new Document();
intrusive_ptr<Document> doc2 =  new Document();
CompareOptions* compareOptions = new CompareOptions();
compareOptions->SetIgnoreFormatting(true);
doc1->Compare(doc2, L"E-iceblue", DateTime::GetNow(), compareOptions);
```

---

# spire.doc cpp word count
## Count characters and words in a Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document =  new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Count the number of words.
wstring content;
content.append(L"CharCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCount()));
content.append(L"\n");
content.append(L"CharCountWithSpace: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCountWithSpace()));
content.append(L"\n");
content.append(L"WordCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetWordCount()));
```

---

# spire.doc cpp document properties
## set built-in document properties
```cpp
// Set document built-in properties
document->GetBuiltinDocumentProperties()->SetTitle(L"Document Demo Document");
document->GetBuiltinDocumentProperties()->SetSubject(L"demo");
document->GetBuiltinDocumentProperties()->SetAuthor(L"James");
document->GetBuiltinDocumentProperties()->SetCompany(L"e-iceblue");
document->GetBuiltinDocumentProperties()->SetManager(L"Jakson");
document->GetBuiltinDocumentProperties()->SetCategory(L"Doc Demos");
document->GetBuiltinDocumentProperties()->SetKeywords(L"Document, Property, Demo");
document->GetBuiltinDocumentProperties()->SetComments(L"This document is just a demo.");
```

---

# spire.doc cpp file operations
## load and save document to disk
```cpp
using namespace Spire::Doc;

int main()
{
	// Create a new document
	intrusive_ptr<Document> doc = new Document();
	// Load the document from the absolute/relative path on disk.
	doc->LoadFromFile(L"Template.docx");

	// Save the document to disk
	doc->SaveToFile(L"LoadAndSaveToDisk.docx", FileFormat::Docx);
	doc->Close();
}
```

---

# spire.doc cpp stream operations
## load document from stream and save to stream
```cpp
// Open the stream. Read only access is enough to load a document.
intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());

// Load the entire document into memory.
intrusive_ptr<Document> doc = new Document(stream);

// You can close the stream now, it is no longer needed because the document is in memory.
stream->Close();
// Do something with the document

// Convert the document to a different format and save to stream.
intrusive_ptr<Stream> newStream = new Stream();
doc->SaveToStream(newStream, FileFormat::Rtf);

newStream->Save(outputFile.c_str());

doc->Close();
```

---

# spire.doc cpp document traversal
## recursively traverse all document objects in a Word document

```cpp
wstring GetDocumentObjectType(DocumentObjectType type)
{
	// Function to convert document object type enum to string representation
	switch (type)
	{
	case DocumentObjectType::Any:
		return L"Any";
	case DocumentObjectType::Body:
		return L"Body";
	case DocumentObjectType::BookmarkEnd:
		return L"BookmarkEnd";
	// ... (other document object types)
	case DocumentObjectType::Undefined:
		return L"Undefined";
	default:
		return L"";
	}
}

void TraverseDocumentObjects(intrusive_ptr<Document> document)
{
	// Traverse all document objects
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		int SectionIndex = document->GetIndex(section);

		int sectionChildObjectsCount = section->GetBody()->GetChildObjects()->GetCount();

		for (int j = 0; j < sectionChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			int objIndex = section->GetBody()->GetIndex(obj);
			DocumentObjectType objType = obj->GetDocumentObjectType();

			if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				intrusive_ptr<Paragraph> paragraph = Object::Dynamic_cast<Paragraph>(obj);
				int paragraphIndex = section->GetBody()->GetIndex(paragraph);

				int paraChildCount = paragraph->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildCount; k++)
				{
					intrusive_ptr<DocumentObject> obj2 = paragraph->GetChildObjects()->GetItem(k);
					int obj2Index = paragraph->GetIndex(obj2);
					DocumentObjectType obj2Type = obj2->GetDocumentObjectType();
				}
			}
		}
	}
}
```

---

# spire.doc cpp document view
## set Word document view modes
```cpp
//Create Word document.
intrusive_ptr<Document> document =  new Document();

//Set Word view modes.
document->GetViewSetup()->SetDocumentViewType(DocumentViewType::WebLayout);
document->GetViewSetup()->SetZoomPercent(150);
document->GetViewSetup()->SetZoomType(ZoomType::None);
```

---

# spire.doc update document properties
## update last saved date of document
```cpp
intrusive_ptr<DateTime> LocalTimeToGreenwishTime(intrusive_ptr<DateTime> lacalTime)
{
	intrusive_ptr<TimeZone> localTimeZone = TimeZone::GetCurrentTimeZone();
	intrusive_ptr<TimeSpan> timeSpan = localTimeZone->GetUtcOffset(lacalTime);

	intrusive_ptr<DateTime> greenwishTime = DateTime::op_Subtraction(lacalTime, timeSpan);
	return greenwishTime;
}

intrusive_ptr<Document> document = new Document();

//Update the last saved date
document->GetBuiltinDocumentProperties()->SetLastSaveDate(LocalTimeToGreenwishTime(DateTime::GetNow()));
document->Close();
```

---

# Spire.Doc C++ Document Operation
## Add a cloned section from one document to another document
```cpp
//Open a Word document as target document
intrusive_ptr<Document> TarDoc = new Document();
//Open a Word document as source document
intrusive_ptr<Document> SouDoc = new Document();

//Get the second section from source document
intrusive_ptr<Section> Ssection = SouDoc->GetSections()->GetItemInSectionCollection(1);

//Add the cloned section in target document
TarDoc->GetSections()->Add(Ssection->CloneSection());
```

---

# spire.doc cpp document operation
## clone a Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Clone the word document.
intrusive_ptr<Document> newDoc = document->CloneDocument();
```

---

# spire.doc cpp document operation
## copy content from one document to another
```cpp
//Copy content from source file and insert them to the target file.
int sectionCount = sourceDoc->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
    intrusive_ptr<Section> sec = sourceDoc->GetSections()->GetItemInSectionCollection(i);
    int sectionChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
    for (int j = 0; j < sectionChildObjectsCount; j++)
    {
        intrusive_ptr<DocumentObject> documentObject = sec->GetBody()->GetChildObjects()->GetItem(j);
        destinationDoc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(documentObject->Clone());
    }
}
```

---

# spire.doc cpp document operation
## append document while keeping same format
```cpp
//Keep same format of source document
srcDoc->SetKeepSameFormat(true);

//Copy the sections of source document to destination document
int sectionCount = srcDoc->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
    intrusive_ptr<Section> section = srcDoc->GetSections()->GetItemInSectionCollection(i);
    destDoc->GetSections()->Add(section->CloneSection());
}
```

---

# spire.doc cpp document operation
## link headers and footers when appending documents
```cpp
using namespace Spire::Doc;

int main() {
	//Load the source file
	intrusive_ptr<Document> srcDoc = new Document();
	
	//Load the destination file
	intrusive_ptr<Document> dstDoc = new Document();

	//Link the headers and footers in the source file
	srcDoc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader()->SetLinkToPrevious(true);
	srcDoc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->SetLinkToPrevious(true);

	//Clone the sections of source to destination
	int sectionCount = srcDoc->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = srcDoc->GetSections()->GetItemInSectionCollection(i);
		dstDoc->GetSections()->Add(section->CloneSection());
	}
}
```

---

# spire.doc cpp merge documents
## merge two word documents into one
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

intrusive_ptr<Document> documentMerge = new Document();

int sectionCount = documentMerge->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
	intrusive_ptr<Section> section = documentMerge->GetSections()->GetItemInSectionCollection(i);
	document->GetSections()->Add(section->CloneSection());
}
```

---

# spire.doc cpp merge documents
## merge multiple documents on the same page
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();
//Clone a destination document
intrusive_ptr<Document> destinationDocument = new Document();

//Traverse sections
int sectionCount = document->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
    intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(i);
    int sectionChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
    for (int j = 0; j < sectionChildObjectsCount; j++)
    {
        intrusive_ptr<DocumentObject> documentObject = sec->GetBody()->GetChildObjects()->GetItem(j);
        destinationDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(documentObject->Clone());
    }
}
```

---

# spire.doc cpp document operation
## preserve theme when appending documents
```cpp
//Create a new Word document
intrusive_ptr<Document> newWord = new Document();

//Clone default style, theme, compatibility from the source document to the destination document
doc->CloneDefaultStyleTo(newWord);
doc->CloneThemesTo(newWord);
doc->CloneCompatibilityTo(newWord);

//Add the cloned section to destination document
newWord->GetSections()->Add(doc->GetSections()->GetItemInSectionCollection(0)->CloneSection());
```

---

# spire.doc cpp section break
## set section break type as continuous in document
```cpp
intrusive_ptr<Document> doc = new Document();

int sectionCount = doc->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
    intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
    //Set section break as continuous
    sec->SetBreakCode(SectionBreakType::NoBreak);
}
```

---

# spire.doc cpp document operation
## insert text from one document into another
```cpp
//Load the Word document
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile_1.c_str());

//Insert document from file
doc->InsertTextFromFile(inputFile_2.c_str(), FileFormat::Auto);
```

---

# Spire.Doc C++ Document Operation
## Split Document by Page Break
```cpp
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
```

---

# spire.doc cpp document operation
## split document by section break
```cpp
// Create Word document.
intrusive_ptr<Document> document = new Document();

// Load the file from disk.
document->LoadFromFile(L"input_file_path.docx");

// Define another new word document object.
intrusive_ptr<Document> newWord;

// Split a Word document into multiple documents by section break.
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    std::wstring result = L"output_path/SplitDocBySectionBreak_" + to_wstring(i) + L".docx";
    newWord = new Document();
    newWord->GetSections()->Add(document->GetSections()->GetItemInSectionCollection(i)->CloneSection());

    // Save to file.
    newWord->SaveToFile(result.c_str());
    newWord->Close();
}
```

---

# spire.doc cpp document operation
## split document into multiple html pages
```cpp
void SplitDocIntoMultipleHtml(const wstring& input, const wstring& outDirectory)
{
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(input.c_str());

	intrusive_ptr<Document> subDoc = nullptr;
	bool first = true;
	int index = 0;
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(i);
		int secChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secChildObjectsCount; j++)
		{
			intrusive_ptr<DocumentObject> element = sec->GetBody()->GetChildObjects()->GetItem(j);
			if (IsInNextDocument(element))
			{
				if (!first)
				{
					//Embed css style and image data into html page
					subDoc->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::Internal);
					subDoc->GetHtmlExportOptions()->SetImageEmbedded(true);
					//Save to html file
					wstring filePath = outDirectory + L"out-" + to_wstring(index++) + L".html";
					subDoc->SaveToFile(filePath.c_str(), FileFormat::Html);
					subDoc = nullptr;
				}
				first = false;
			}
			if (subDoc == nullptr)
			{
				subDoc = new Document();
				subDoc->AddSection();
			}
			subDoc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add(element->Clone());
		}
	}
	if (subDoc != nullptr)
	{
		//Embed css style and image data into html page
		subDoc->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::Internal);
		subDoc->GetHtmlExportOptions()->SetImageEmbedded(true);
		//Save to html file
		wstring filePath = outDirectory + L"out-" + to_wstring(index++) + L".html";
		subDoc->SaveToFile(filePath.c_str(), FileFormat::Html);
		subDoc->Close();
	}
}

bool IsInNextDocument(intrusive_ptr<DocumentObject> element)
{
	if (Object::CheckType<Paragraph>(element))
	{
		intrusive_ptr<Paragraph> p = boost::dynamic_pointer_cast<Paragraph>(element);
		if (wcscmp(p->GetStyleName(), L"Heading1") == 0)
		{
			return true;
		}
	}
	return false;
}
```

---

# spire.doc cpp track changes
## accept or reject tracked changes in word document
```cpp
//Get the first section and the paragraph we want to accept/reject the changes.
intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(0);

//Accept the changes or reject the changes.
para->GetDocument()->AcceptChanges();
//para.Document.RejectChanges();
```

---

# spire.doc cpp track changes
## enable track changes in document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Enable the track changes.
document->SetTrackChanges(true);
```

---

# spire.doc cpp document revisions
## extract insert and delete revisions from document
```cpp
intrusive_ptr<Document> document = new Document();

//Traverse sections
int sectionCount = document->GetSections()->GetCount();
for (int i = 0; i < sectionCount; i++)
{
    intrusive_ptr<Section> sec = document->GetSections()->GetItemInSectionCollection(i);
    //Iterate through the element under GetBody() in the section
    int secChildObjectsCount = sec->GetBody()->GetChildObjects()->GetCount();
    for (int j = 0; j < secChildObjectsCount; j++)
    {
        intrusive_ptr<DocumentObject> docItem = sec->GetBody()->GetChildObjects()->GetItem(j);
        if (Object::CheckType<Paragraph>(docItem))
        {
            intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(docItem);
            //Determine if the paragraph is an insertion revision
            if (para->GetIsInsertRevision())
            {
                //Get insertion revision
                intrusive_ptr<EditRevision> insRevison = para->GetInsertRevision();
                
                //Get insertion revision type
                EditRevisionType insType = insRevison->GetType();
                
                //Get insertion revision author
                std::wstring insAuthor = insRevison->GetAuthor();
            }
            //Determine if the paragraph is a delete revision
            else if (para->GetIsDeleteRevision())
            {
                intrusive_ptr<EditRevision> delRevison = para->GetDeleteRevision();
                EditRevisionType delType = delRevison->GetType();
                std::wstring delAuthor = delRevison->GetAuthor();
            }
            //Iterate through the element in the paragraph
            int paraChildObjectsCount = para->GetChildObjects()->GetCount();
            for (int k = 0; k < paraChildObjectsCount; k++)
            {
                intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
                if (Object::CheckType<TextRange>(obj))
                {
                    intrusive_ptr<TextRange> textRange = boost::dynamic_pointer_cast<TextRange>(obj);
                    //Determine if the textrange is an insertion revision
                    if (textRange->GetIsInsertRevision())
                    {
                        intrusive_ptr<EditRevision> insRevison = textRange->GetInsertRevision();
                        EditRevisionType insType = insRevison->GetType();
                        std::wstring insAuthor = insRevison->GetAuthor();
                    }
                    else if (textRange->GetIsDeleteRevision())
                    {
                        //Determine if the textrange is a delete revision
                        intrusive_ptr<EditRevision> delRevison = textRange->GetDeleteRevision();
                        EditRevisionType delType = delRevison->GetType();
                        std::wstring delAuthor = delRevison->GetAuthor();
                    }
                }
            }
        }
    }
}

wstring getRevisionType(EditRevisionType type)
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
}
```

---

# spire.doc cpp variables
## add document variables to Word document
```cpp
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
```

---

# spire.doc cpp variables
## count document variables
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Get the number of variables in the document.
int number = document->GetVariables()->GetCount();
```

---

# spire.doc cpp variables
## get document variables
```cpp
intrusive_ptr<Document> document = new Document();
//Load the file from disk.
document->LoadFromFile(inputFile.c_str());
wstring stringBuilder;

stringBuilder.append(L"This document has following variables:\n");
int variablesCount = document->GetVariables()->GetCount();
for (int i = 0; i < variablesCount; i++)
{
	std::wstring name = document->GetVariables()->GetNameByIndex(i);
	std::wstring value = document->GetVariables()->GetValueByIndex(i);
	stringBuilder.append(L"Name: " + name + L", " + L"Value: " + value);
	stringBuilder.append(L"\n");
}
```

---

# spire.doc cpp variables
## remove variables from document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Remove the variable by name.
document->GetVariables()->Remove(L"A1");
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp variables
## retrieve document variables by index and name
```cpp
//Retrieve name of the variable by index.
std::wstring s1 = document->GetVariables()->GetNameByIndex(0);

//Retrieve value of the variable by index.
std::wstring s2 = document->GetVariables()->GetValueByIndex(0);

//Retrieve the value of the variable by name.
std::wstring s3 = document->GetVariables()->GetItem(L"A1");
```

---

# spire.doc cpp gradient background
## set gradient background for document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Set the background type as Gradient.
document->GetBackground()->SetType(BackgroundType::Gradient);
intrusive_ptr<BackgroundGradient> Test = document->GetBackground()->GetGradient();

//Set the first color and second color for Gradient.
Test->SetColor1(Color::GetWhite());
Test->SetColor2(Color::GetLightBlue());

//Set the Shading style and Variant for the gradient.
Test->SetShadingVariant(GradientShadingVariant::ShadingDown);
Test->SetShadingStyle(GradientShadingStyle::Horizontal);
```

---

# spire.doc cpp background
## set background type as picture for document
```cpp
using namespace Spire::Doc;

intrusive_ptr<Document> document = new Document();
//set the background type as picture.
document->GetBackground()->SetType(BackgroundType::Picture);
```

---

# spire.doc cpp page setup
## add gutter to word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Get the first section of the document
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

//Set gutter
section->GetPageSetup()->SetGutter(100.0f);
```

---

# spire.doc cpp page setup
## add line numbers to word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Set the start value of the line numbers.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingStartValue(1);

//Set the interval between displayed numbers.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingStep(6);

//Set the distance between line numbers and text.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingDistanceFromText(40.0f);

//Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->SetLineNumberingRestartMode(LineNumberingRestartMode::Continuous);
```

---

# spire.doc cpp page setup
## add page borders to document
```cpp
//Define the border style.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetBorderType(BorderStyle::DotDash);

//Define the border color.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetColor(Color::GetRed());

//Set the line width.
document->GetSections()->GetItemInSectionCollection(0)->GetPageSetup()->GetBorders()->SetLineWidth(2);
```

---

# spire.doc cpp page setup
## add page numbers in document sections
```cpp
// Repeat step2 and Step3 for the rest sections, so change the code with for loop.
for (int i = 0; i < 3; i++)
{
    intrusive_ptr<HeaderFooter> footer = document->GetSections()->GetItemInSectionCollection(i)->GetHeadersFooters()->GetFooter();
    intrusive_ptr<Paragraph> footerParagraph = footer->AddParagraph();
    footerParagraph->AppendField(L"page number", FieldType::FieldPage);
    footerParagraph->AppendText(L" of ");
    footerParagraph->AppendField(L"number of pages", FieldType::FieldSectionPages);
    footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

    if (i == 2)
    {
        break;
    }
    else
    {
        document->GetSections()->GetItemInSectionCollection(i + 1)->GetPageSetup()->SetRestartPageNumbering(true);
        document->GetSections()->GetItemInSectionCollection(i + 1)->GetPageSetup()->SetPageStartingNumber(1);
    }
}
```

---

# spire.doc cpp page setup
## configure different page setup for document sections
```cpp
//Create a Word document
intrusive_ptr<Document> doc = new Document();

//Get the second section 
intrusive_ptr<Section> SectionTwo = doc->GetSections()->GetItemInSectionCollection(1);

//Set the orientation
SectionTwo->GetPageSetup()->SetOrientation(PageOrientation::Landscape);

//Set page size
//SectionTwo.PageSetup.PageSize = new SizeF(800, 800);
```

---

# spire.doc cpp page break
## insert section break in word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

intrusive_ptr<Section> section = document->AddSection();

//insert a break code
section = document->AddSection();
section->AddParagraph()->InsertSectionBreak(SectionBreakType::NewPage);
```

---

# spire.doc cpp page break
## insert page break after specific text
```cpp
//Find the specified word "technology" where we want to insert the page break.
std::vector<intrusive_ptr<TextSelection>> selections = document->FindAllString(L"technology", true, true);

//Traverse each word "technology".
for (intrusive_ptr<TextSelection> ts : selections)
{
    intrusive_ptr<TextRange> range = ts->GetAsOneRange();
    intrusive_ptr<Paragraph> paragraph = range->GetOwnerParagraph();
    int index = paragraph->GetChildObjects()->IndexOf(range);

    //Create a new instance of page break and insert a page break after the word "technology".
    intrusive_ptr<Break> pageBreak = new Break(document, BreakType::PageBreak);
    paragraph->GetChildObjects()->Insert(index + 1, pageBreak);

    //C# TO C++ CONVERTER TODO TASK: A 'delete pageBreak' statement was not added since pageBreak was passed to a method or constructor. Handle memory management manually.
}
```

---

# spire.doc cpp page break
## insert page break using second approach
```cpp
using namespace Spire::Doc;

// Create Word document.
intrusive_ptr<Document> document = new Document();

// Insert page break.
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(3)->AppendBreak(BreakType::PageBreak);

document->Close();
```

---

# spire.doc cpp section break
## insert section break in document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Insert section break. There are five section break options including EvenPage, NewColumn, NewPage, NoBreak, OddPage.
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(1)->InsertSectionBreak(SectionBreakType::NoBreak);
```

---

# Spire.Doc C++ Page Setup
## Configure document page settings including margins, headers, and footers
```cpp
using namespace Spire::Doc;

void InsertHeaderAndFooter(intrusive_ptr<Section> section);

int main() {
    // Create Word document
    intrusive_ptr<Document> document = new Document();
    intrusive_ptr<Section> section = document->AddSection();

    // Set page size and margins (unit: point, 1point = 0.3528 mm)
    section->GetPageSetup()->SetPageSize(PageSize::A4());
    section->GetPageSetup()->GetMargins()->SetTop(72.0f);
    section->GetPageSetup()->GetMargins()->SetBottom(72.0f);
    section->GetPageSetup()->GetMargins()->SetLeft(89.85f);
    section->GetPageSetup()->GetMargins()->SetRight(89.85f);

    // Insert header and footer
    InsertHeaderAndFooter(section);
}

void InsertHeaderAndFooter(intrusive_ptr<Section> section)
{
    intrusive_ptr<HeaderFooter> header = section->GetHeadersFooters()->GetHeader();
    intrusive_ptr<HeaderFooter> footer = section->GetHeadersFooters()->GetFooter();

    // Setup header with text and image
    intrusive_ptr<Paragraph> headerParagraph = header->AddParagraph();
    intrusive_ptr<DocPicture> headerPicture = headerParagraph->AppendPicture();
    intrusive_ptr<TextRange> headerText = headerParagraph->AppendText(L"Demo of Spire.Doc");
    
    // Format header
    headerText->GetCharacterFormat()->SetFontName(L"Arial");
    headerText->GetCharacterFormat()->SetFontSize(10);
    headerText->GetCharacterFormat()->SetItalic(true);
    headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
    headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
    
    // Configure header image layout
    headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
    headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
    headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
    headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
    headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

    // Setup footer with image and page numbers
    intrusive_ptr<Paragraph> footerParagraph = footer->AddParagraph();
    intrusive_ptr<DocPicture> footerPicture = footerParagraph->AppendPicture();
    
    // Configure footer image layout
    footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
    footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
    footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
    footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
    footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);
    
    // Add page numbering
    footerParagraph->AppendField(L"page number", FieldType::FieldPage);
    footerParagraph->AppendText(L" of ");
    footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
    footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
    
    // Add footer border
    footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
}
```

---

# spire.doc cpp page setup
## remove page breaks from document
```cpp
//Traverse every paragraph of the first section of the document.
for (int j = 0; j < document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetCount(); j++)
{
    intrusive_ptr<Paragraph> p = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(j);

    //Traverse every child object of a paragraph.
    for (int i = 0; i < p->GetChildObjects()->GetCount(); i++)
    {
        intrusive_ptr<DocumentObject> obj = p->GetChildObjects()->GetItem(i);

        //Find the page break object.
        if (obj->GetDocumentObjectType() == DocumentObjectType::Break)
        {
            intrusive_ptr<Break> b = Object::Dynamic_cast<Break>(obj);

            //Remove the page break object from paragraph.
            p->GetChildObjects()->Remove(b);
        }
    }
}
```

---

# spire.doc cpp page setup
## reset page numbering in document sections
```cpp
//Use section method to combine all documents into one word document.
for (int i = 0; i < document2->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> sec = document2->GetSections()->GetItemInSectionCollection(i);
    document1->GetSections()->Add(sec->CloneSection());
}
for (int j = 0; j < document3->GetSections()->GetCount(); j++)
{
    intrusive_ptr<Section> sec = document3->GetSections()->GetItemInSectionCollection(j);
    document1->GetSections()->Add(sec->CloneSection());
}

//Traverse every section of document1.
for (int k = 0; k < document1->GetSections()->GetCount(); k++)
{
    intrusive_ptr<Section> sec = document1->GetSections()->GetItemInSectionCollection(k);
    //Traverse every object of the footer.
    for (int m = 0; m < sec->GetHeadersFooters()->GetFooter()->GetChildObjects()->GetCount(); m++)
    {
        intrusive_ptr<DocumentObject> obj = sec->GetHeadersFooters()->GetFooter()->GetChildObjects()->GetItem(m);
        if (obj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
        {
            intrusive_ptr<DocumentObject> para = obj->GetChildObjects()->GetItem(m);
            for (int n = 0; n < para->GetChildObjects()->GetCount(); n++)
            {
                intrusive_ptr<DocumentObject> item = para->GetChildObjects()->GetItem(n);
                if (item->GetDocumentObjectType() == DocumentObjectType::Field)
                {
                    //Find the item and its field type is FieldNumPages.
                    if ((Object::Dynamic_cast<Field>(item))->GetType() == FieldType::FieldNumPages)
                    {
                        //Change field type to FieldSectionPages.
                        (Object::Dynamic_cast<Field>(item))->SetType(FieldType::FieldSectionPages);
                    }
                }
            }
        }
    }
}

//Restart page number of section and set the starting page number to 1.
document1->GetSections()->GetItemInSectionCollection(1)->GetPageSetup()->SetRestartPageNumbering(true);
document1->GetSections()->GetItemInSectionCollection(1)->GetPageSetup()->SetPageStartingNumber(1);

document1->GetSections()->GetItemInSectionCollection(2)->GetPageSetup()->SetRestartPageNumbering(true);
document1->GetSections()->GetItemInSectionCollection(2)->GetPageSetup()->SetPageStartingNumber(1);
```

---

# spire.doc cpp conversion
## convert document to byte array and back
```cpp
// Save document to stream
intrusive_ptr<Stream> outStream = new Stream();
document->SaveToStream(outStream, FileFormat::Docx);

// Convert the document to bytes
std::vector<byte> docBytes = outStream->ToArray();

// Convert bytes back to document
intrusive_ptr<Stream> inStream = new Stream(docBytes.data(), docBytes.size());
intrusive_ptr<Document> restoredDocument = new Document(inStream);
```

---

# spire.doc cpp conversion
## convert HTML to image
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

//Save to image in the default format of png.
intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
```

---

# spire.doc cpp html to pdf conversion
## convert HTML file to PDF format using Spire.Doc library
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load HTML file from disk.
document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

//Save to PDF file.
document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
document->Close();
```

---

# spire.doc cpp conversion
## convert HTML to XML
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the HTML file from disk.
document->LoadFromFile(inputFile.c_str());

//Save to XML file.
document->SaveToFile(outputFile.c_str(), FileFormat::Xml);
document->Close();
```

---

# spire.doc cpp conversion
## convert HTML to XPS format
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

//Save to file.
document->SaveToFile(outputFile.c_str(), FileFormat::XPS);
document->Close();
```

---

# spire.doc cpp conversion
## convert image to pdf
```cpp
//Create a new document
intrusive_ptr<Document>  doc = new Document();
//Create a new section
intrusive_ptr<Section> section = doc->AddSection();
//Create a new paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
//Add a picture for paragraph
intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(inputFile.c_str());
//Set the page size to the same size as picture
section->GetPageSetup()->SetPageSize(new SizeF(picture->GetWidth(), picture->GetHeight()));
//Set A4 page size
section->GetPageSetup()->SetPageSize(PageSize::A4());
//Set the page margins
section->GetPageSetup()->GetMargins()->SetTop(10.0f);
section->GetPageSetup()->GetMargins()->SetLeft(25.0f);

doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
doc->Close();
```

---

# spire.doc cpp conversion
## convert ODT file to Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the ODT file from disk.
document->LoadFromFile(inputFile.c_str());

//Save to Word file.
document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
document->Close();
```

---

# RTF to HTML Conversion
## Convert RTF document to HTML format using Spire.Doc library
```cpp
using namespace Spire::Doc;

//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Save to file.
document->SaveToFile(outputFile.c_str(), FileFormat::Html);
document->Close();
```

---

# spire.doc cpp conversion
## convert RTF file to PDF format
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Save to file.
document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to image
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
//Obtain image data in the default format of png,you can use it to convert other image format
std::vector<byte> imgBytes = imageStream->ToArray();
std::ofstream outFile(outputFile, std::ios::binary);
if (outFile.is_open())
{
    outFile.write(reinterpret_cast<const char*>(imgBytes.data()), imgBytes.size());
    outFile.close();
}
document->Close();
imageStream->Dispose();
```

---

# spire.doc cpp conversion
## convert Word document to ODT format
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Save doc file.
document->SaveToFile(outputFile.c_str(), FileFormat::Odt);
document->Close();
```

---

# spire.doc cpp conversion
## convert document to PCL format
```cpp
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());
doc->SaveToFile(outputFile.c_str(), FileFormat::PCL);
doc->Close();
```

---

# spire.doc cpp conversion
## convert document to PostScript format
```cpp
//Create Word document.
intrusive_ptr<Document> doc = new Document();

//Load the file from disk.
doc->LoadFromFile(inputFile.c_str());
//Save to PostScript file.
doc->SaveToFile(outputFile.c_str(), FileFormat::PostScript);
doc->Close();
```

---

# spire.doc cpp conversion
## convert doc to rtf
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Save doc file.
document->SaveToFile(outputFile.c_str(), FileFormat::Rtf);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to SVG format
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());
document->SaveToFile(outputFile.c_str(), FileFormat::SVG);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to XML format
```cpp
wstring inputFile = L"Summary_of_Science.doc";
wstring outputFile = L"ToXML.xml";

//Create word document.
intrusive_ptr<Document> document = new Document();
//Load file from disk
document->LoadFromFile(inputFile.c_str());
//Save the document to a xml file.
document->SaveToFile(outputFile.c_str(), FileFormat::Xml);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to XPS format
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Save the document to a xps file.
document->SaveToFile(outputFile.c_str(), FileFormat::XPS);
document->Close();
```

---

# spire.doc cpp conversion
## convert text file to word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Save the file.
document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
document->Close();
```

---

# Word to PDF/A Conversion
## Convert a Word document to PDF/A format with PDF_A1B conformance level
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Set the Conformance-level of the Pdf file to PDF_A1B.
intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();
toPdf->SetPdfConformanceLevel(PdfConformanceLevel::Pdf_A1B);

//Save the file.
document->SaveToFile(outputFile.c_str(), toPdf);
document->Close();
```

---

# Spire.Doc C++ Conversion
## Convert Word document to text file
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the Word document from disk.
document->LoadFromFile(inputFilePath);

//Save as text file.
document->SaveToFile(outputFilePath, FileFormat::Txt);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to WordXML format
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//For word 2003:
document->SaveToFile(outputFile_2003.c_str(), FileFormat::WordML);

//For word 2007:
document->SaveToFile(outputFile_2007.c_str(), FileFormat::WordXml);
document->Close();
```

---

# spire.doc cpp conversion
## convert HTML file to Word document
```cpp
//Create a Document object
intrusive_ptr<Document> document = new Document();

//Load HTML file
document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

//Save as Word document
document->SaveToFile(outputFile.c_str(), FileFormat::Docx);

//Close the document
document->Close();
```

---

# spire.doc cpp html to word conversion
## convert HTML string to Word document
```cpp
// Get HTML string.
ifstream in(inputFile.c_str(), ios::in);
istreambuf_iterator<char> begin(in), end;
wstring HTML(begin, end);
in.close();

// Create a new document.
intrusive_ptr<Document> document = new Document();

// Add a section.
intrusive_ptr<Section> sec = document->AddSection();

// Add a paragraph and append HTML string.
intrusive_ptr<Paragraph> para = sec->AddParagraph();
para->AppendHTML(HTML.c_str());

// Save it to a Word file.
document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
document->Close();
```

---

# spire.doc cpp epub conversion
## add cover image to EPUB document
```cpp
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
picture->LoadImageSpire((input_path + L"Cover.png").c_str());
doc->SaveToEpub(outputFile.c_str(), picture);
doc->Close();
```

---

# spire.doc cpp conversion
## convert Word document to EPUB format
```cpp
//Create a new document.
intrusive_ptr<Document> doc = new Document();
//Load document from file
doc->LoadFromFile(inputFile.c_str());
//Save the document to a Epub file.
doc->SaveToFile(outputFile.c_str(), FileFormat::EPub);
doc->Close();
```

---

# spire.doc cpp conversion
## convert Word document to HTML format
```cpp
//Create word document
intrusive_ptr<Document>  document = new Document();
document->LoadFromFile(inputFile.c_str());

//Save doc file.
document->SaveToFile(outputFile.c_str(), FileFormat::Html);
document->Close();
```

---

# spire.doc cpp html export
## configure HTML export options for Word document conversion
```cpp
//Open a Word document.
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Set whether the css styles are embeded or not. 
document->GetHtmlExportOptions()->SetCssStyleSheetFileName(L"sample.css");
document->GetHtmlExportOptions()->SetCssStyleSheetType(CssStyleSheetType::External);

//Set whether the images are embeded or not. 
document->GetHtmlExportOptions()->SetImageEmbedded(false);
document->GetHtmlExportOptions()->SetImagesPath(output_path.c_str());

//Set the option whether to export form fields as plain text or not.
document->GetHtmlExportOptions()->SetIsTextInputFormFieldAsText(true);

//Save the document to a html file.
document->SaveToFile(outputFile.c_str(), FileFormat::Html);
document->Close();
```

---

# spire.doc cpp pdf conversion
## disable hyperlinks when converting document to pdf
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Create an instance of ToPdfParameterList.
intrusive_ptr<ToPdfParameterList> pdf = new ToPdfParameterList();

//Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
//Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
pdf->SetDisableLink(true);
```

---

# spire.doc cpp conversion
## convert Word document to PDF with all fonts embedded
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());
//embeds full fonts by default when IsEmbeddedAllFonts is set to true.
intrusive_ptr<ToPdfParameterList> ppl = new ToPdfParameterList();
ppl->SetIsEmbeddedAllFonts(true);

//Save doc file to pdf.
document->SaveToFile(outputFile.c_str(), ppl);
document->Close();
```

---

# spire.doc cpp pdf conversion
## embed non-installed fonts when converting to PDF
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Embed the non-installed fonts.
intrusive_ptr<ToPdfParameterList> parms = new ToPdfParameterList();
std::vector<intrusive_ptr<PrivateFontPath>> fonts;

intrusive_ptr<PrivateFontPath> tempVar = new PrivateFontPath(L"PT Serif Caption", (input_path + L"PT_Serif-Caption-Web-Regular.ttf").c_str());
fonts.push_back(tempVar);
parms->SetPrivateFontPaths(fonts);

//Save doc file to pdf.
document->SaveToFile(outputFile.c_str(), parms);
document->Close();
```

---

# spire.doc cpp pdf conversion
## keep hidden text when converting to PDF
```cpp
//When convert to PDF file, set the property IsHidden as true.
intrusive_ptr<ToPdfParameterList> pdf = new ToPdfParameterList();
pdf->SetIsHidden(true);
```

---

# spire.doc cpp image quality
## set JPEG quality for document conversion
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
document->SetJPEGQuality(40);
```

---

# spire.doc cpp font embedding
## specify embedded font when converting Word to PDF
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

document->LoadFromFile(inputFile.c_str());
//Specify embedded font
intrusive_ptr<ToPdfParameterList> parms = new ToPdfParameterList();
std::vector<std::wstring> part;
part.push_back(L"PT Serif Caption");
parms->SetEmbeddedFontNameList(part);
document->SaveToFile(outputFile.c_str(), parms);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to PDF
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Save the document to a PDF file.
document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
document->Close();
```

---

# spire.doc cpp conversion
## convert Word document to PDF with bookmarks
```cpp
//Create Word document
intrusive_ptr<Document> document = new Document();
//Load the document from disk
document->LoadFromFile(inputFile.c_str());
intrusive_ptr<ToPdfParameterList> parames = new ToPdfParameterList();
//Set CreateWordBookmarks to true
parames->SetCreateWordBookmarks(true);

//Create bookmarks using Headings
//parames->SetCreateWordBookmarksUsingHeadings(true);

//Create bookmarks using word bookmarks
parames->SetCreateWordBookmarksUsingHeadings(false);
document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
document->Close();
```

---

# spire.doc cpp pdf conversion
## convert Word document to PDF with password protection
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//create a parameter
intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();

//set the password
std::wstring password = L"E-iceblue";
toPdf->GetPdfSecurity()->Encrypt(L"password", password.c_str(), PdfPermissionsFlags::Default, PdfEncryptionKeySize::Key128Bit);
//save doc file.
document->SaveToFile(outputFile.c_str(), toPdf);
document->Close();
```

---

# spire.doc cpp font color
## change font color in document styles
```cpp
//Initialize a document
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> sec = document->AddSection();

//Add default title style to document and modify
intrusive_ptr<Style> titleStyle = document->AddStyle(BuiltinStyle::Title);
titleStyle->GetCharacterFormat()->SetFontName(L"cambria");
titleStyle->GetCharacterFormat()->SetFontSize(28);
titleStyle->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));

//Judge if it is Paragraph Style and then set paragraph format
if (Object::Dynamic_cast<ParagraphStyle>(titleStyle) != nullptr)
{
	intrusive_ptr<ParagraphStyle> ps = Object::Dynamic_cast<ParagraphStyle>(titleStyle);
	ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetColor(Color::FromArgb(42, 123, 136));
	ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetLineWidth(1.5f);
	ps->GetParagraphFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
}

//Add default heading1 style
intrusive_ptr<Style> heading1Style = document->AddStyle(BuiltinStyle::Heading1);
heading1Style->GetCharacterFormat()->SetFontName(L"cambria");
heading1Style->GetCharacterFormat()->SetFontSize(14);
heading1Style->GetCharacterFormat()->SetBold(true);
heading1Style->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));
```

---

# spire.doc cpp font embedding
## embed private font into word document
```cpp
//Load the document
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());

//Get the first section and add a paragraph
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> p = section->AddParagraph();

//Append text to the paragraph, then set the font name and font size
intrusive_ptr<TextRange> range = p->AppendText(L"Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.");
range->GetCharacterFormat()->SetFontName(L"PT Serif Caption");
range->GetCharacterFormat()->SetFontSize(20);

//Allow embedding font in document
doc->SetEmbedFontsInFile(true);

//Embed private font from font file into the document
intrusive_ptr<PrivateFontPath> tempVar = new PrivateFontPath(L"PT Serif Caption", (input_path + L"PT_Serif-Caption-Web-Regular.ttf").c_str());
doc->GetPrivateFontList().push_back(tempVar);

//Save and launch document
doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
doc->Close();
```

---

# spire.doc cpp font
## set font for text in document
```cpp
//Get the first section 
intrusive_ptr<Section> s = doc->GetSections()->GetItemInSectionCollection(0);

//Get the second paragraph
intrusive_ptr<Paragraph> p = s->GetParagraphs()->GetItemInParagraphCollection(1);

//Create a characterFormat object
intrusive_ptr<CharacterFormat> format = new CharacterFormat(doc);
//Set font
format->SetFontName(L"Arial");
format->SetFontSize(16);

//Loop through the childObjects of paragraph 
int pChildObjectsCount = p->GetChildObjects()->GetCount();
for (int i = 0; i < pChildObjectsCount; i++)
{
	intrusive_ptr<DocumentObject> childObj = p->GetChildObjects()->GetItem(i);
	if (Object::CheckType<TextRange>(childObj))
	{
		//Apply character format
		intrusive_ptr<TextRange> tr = boost::dynamic_pointer_cast<TextRange>(childObj);
		tr->ApplyCharacterFormat(format);
	}
}
```

---

# Spire.Doc C++ ASCII Characters Bullet Style
## Create list styles with different ASCII characters as bullet points and apply them to paragraphs
```cpp
using namespace Spire::Doc;

int main() {
	//Create a new document
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();

	//Create four list styles based on different ASCII characters
	intrusive_ptr<ListStyle> listStyle1 = new ListStyle(document, ListType::Bulleted);
	listStyle1->SetName(L"liststyle");
	listStyle1->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x006e");
	listStyle1->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle1);
	intrusive_ptr<ListStyle> listStyle2 = new ListStyle(document, ListType::Bulleted);
	listStyle2->SetName(L"liststyle2");
	listStyle2->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x0075");
	listStyle2->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle2);
	intrusive_ptr<ListStyle> listStyle3 = new ListStyle(document, ListType::Bulleted);
	listStyle3->SetName(L"liststyle3");
	listStyle3->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00b2");
	listStyle3->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle3);
	intrusive_ptr<ListStyle> listStyle4 = new ListStyle(document, ListType::Bulleted);
	listStyle4->SetName(L"liststyle4");
	listStyle4->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00d8");
	listStyle4->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle4);

	//Add four paragraphs and apply list style separately
	intrusive_ptr<Paragraph> p1 = section->GetBody()->AddParagraph();
	p1->AppendText(L"Spire.Doc for .NET");
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());
	intrusive_ptr<Paragraph> p2 = section->GetBody()->AddParagraph();
	p2->AppendText(L"Spire.Doc for .NET");
	p2->GetListFormat()->ApplyStyle(listStyle2->GetName());
	intrusive_ptr<Paragraph> p3 = section->GetBody()->AddParagraph();
	p3->AppendText(L"Spire.Doc for .NET");
	p3->GetListFormat()->ApplyStyle(listStyle3->GetName());
	intrusive_ptr<Paragraph> p4 = section->GetBody()->AddParagraph();
	p4->AppendText(L"Spire.Doc for .NET");
	p4->GetListFormat()->ApplyStyle(listStyle4->GetName());
}
```

---

# spire.doc cpp character formatting
## demonstrates various character formatting options in Spire.Doc for C++
```cpp
//Initialize a document
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> sec = document->AddSection();
intrusive_ptr<Paragraph> titleParagraph = sec->AddParagraph();
titleParagraph->AppendText(L"Font Styles and Effects ");
titleParagraph->ApplyStyle(BuiltinStyle::Title);

intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
intrusive_ptr<TextRange> tr = paragraph->AppendText(L"Strikethough Text");
tr->GetCharacterFormat()->SetIsStrikeout(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Shadow Text");
tr->GetCharacterFormat()->SetIsShadow(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Small caps Text");
tr->GetCharacterFormat()->SetIsSmallCaps(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Double Strikethough Text");
tr->GetCharacterFormat()->SetDoubleStrike(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Outline Text");
tr->GetCharacterFormat()->SetIsOutLine(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"AllCaps Text");
tr->GetCharacterFormat()->SetAllCaps(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Text");
tr = paragraph->AppendText(L"SubScript");
tr->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SubScript);

tr = paragraph->AppendText(L"And");
tr = paragraph->AppendText(L"SuperScript");
tr->GetCharacterFormat()->SetSubSuperScript(SubSuperScript::SuperScript);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Emboss Text");
tr->GetCharacterFormat()->SetEmboss(true);
tr->GetCharacterFormat()->SetTextColor(Color::GetWhite());

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Hidden:");
tr = paragraph->AppendText(L"Hidden Text");
tr->GetCharacterFormat()->SetHidden(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Engrave Text");
tr->GetCharacterFormat()->SetEngrave(true);
tr->GetCharacterFormat()->SetTextColor(Color::GetWhite());

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"WesternFonts中文字体");
tr->GetCharacterFormat()->SetFontNameAscii(L"Calibri");
tr->GetCharacterFormat()->SetFontNameNonFarEast(L"Calibri");
tr->GetCharacterFormat()->SetFontNameFarEast(L"Simsun");

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Font Size");
tr->GetCharacterFormat()->SetFontSize(20);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Font Color");
tr->GetCharacterFormat()->SetTextColor(Color::GetRed());

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Bold Italic Text");
tr->GetCharacterFormat()->SetBold(true);
tr->GetCharacterFormat()->SetItalic(true);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Underline Style");
tr->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Highlight Text");
tr->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Text has shading");
tr->GetCharacterFormat()->SetTextBackgroundColor(Color::GetGreen());

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Border Around Text");
tr->GetCharacterFormat()->GetBorder()->SetBorderType(BorderStyle::Single);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Text Scale");
tr->GetCharacterFormat()->SetTextScale(150);

paragraph->AppendBreak(BreakType::LineBreak);
tr = paragraph->AppendText(L"Character Spacing is 2 point");
tr->GetCharacterFormat()->SetCharacterSpacing(2);
```

---

# spire.doc cpp styles
## copy document styles between documents
```cpp
//Get the style collections of source document
intrusive_ptr<StyleCollection> styles = srcDoc->GetStyles();

//Copy each style from source to destination document
for (int i = 0; i < styles->GetCount(); i++)
{
    intrusive_ptr<IStyle> style = styles->GetItem(i);
    intrusive_ptr<Style> destStyle = Object::Dynamic_cast<Style>(style);
    destDoc->GetStyles()->Add(destStyle);
}
```

---

# spire.doc cpp character spacing
## get character spacing from document
```cpp
//Get the first section of document
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

//Get the first paragraph 
intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(0);

//Define two variables
wstring fontName = L"";
float fontSpacing = 0;

//Traverse the ChildObjects 
for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
{
    intrusive_ptr<DocumentObject> docObj = paragraph->GetChildObjects()->GetItem(i);
    //If it is TextRange
    if (Object::CheckType<TextRange>(docObj))
    {
        intrusive_ptr<TextRange> textRange = boost::dynamic_pointer_cast<TextRange>(docObj);

        //Get the font name
        fontName = textRange->GetCharacterFormat()->GetFontName();

        //Get the character spacing
        fontSpacing = textRange->GetCharacterFormat()->GetCharacterSpacing();
    }
}
```

---

# spire.doc cpp style
## get text by style name
```cpp
//Create string builder
wstring* builder = new wstring();
//Loop through sections
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    //Loop through paragraphs
    for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(j);
        //Find the paragraph whose style name is "Heading1"
        wstring style_name = para->GetStyleName();
        if (style_name.compare(L"Heading1") == 0)
        {
            //Write the text of paragraph
            builder->append(para->GetText());
            builder->append(L"\n");
        }
    }
}
```

---

# spire.doc cpp lists
## create and apply list styles in word document
```cpp
//Initialize a document
intrusive_ptr<Document> document = new Document();
//Add a section
intrusive_ptr<Section> sec = document->AddSection();
//Add paragraph and set list style
intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
paragraph->AppendText(L"Lists");
paragraph->ApplyStyle(BuiltinStyle::Title);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Numbered List:")->GetCharacterFormat()->SetBold(true);

//Create list style
intrusive_ptr<ListStyle> numberList = new ListStyle(document, ListType::Numbered);
numberList->SetName(L"numberList");
//%1-%9
numberList->GetLevels()->GetItem(1)->SetNumberPrefix(L"%1.");
numberList->GetLevels()->GetItem(1)->SetPatternType(ListPatternType::Arabic);
numberList->GetLevels()->GetItem(2)->SetNumberPrefix(L"%1.%2.");
numberList->GetLevels()->GetItem(2)->SetPatternType(ListPatternType::Arabic);

intrusive_ptr<ListStyle> bulletList = new ListStyle(document, ListType::Bulleted);
bulletList->SetName(L"bulletList");

//add the list style into document
document->GetListStyles()->Add(numberList);
document->GetListStyles()->Add(bulletList);

//Add paragraph and apply the list style
paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 1");
paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 2");
paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 2.1");
paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
paragraph->GetListFormat()->SetListLevelNumber(1);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Bulleted List:")->GetCharacterFormat()->SetBold(true);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 1");
paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());

paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 2");
paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());

paragraph = sec->AddParagraph();
paragraph->AppendText(L"List Item 2.1");
paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());
paragraph->GetListFormat()->SetListLevelNumber(1);
```

---

# spire.doc cpp styles
## apply multiple styles in a paragraph
```cpp
//Create a Word document
intrusive_ptr<Document> doc = new Document();

//Add a section
intrusive_ptr<Section> section = doc->AddSection();

//Add a paragraph
intrusive_ptr<Paragraph> para = section->AddParagraph();

//Add a text range 1 and set its style
intrusive_ptr<TextRange> range = para->AppendText(L"Spire.Doc for .NET ");
range->GetCharacterFormat()->SetFontName(L"Calibri");
range->GetCharacterFormat()->SetFontSize(16.0f);
range->GetCharacterFormat()->SetTextColor(Color::GetBlue());
range->GetCharacterFormat()->SetBold(true);
range->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

//Add a text range 2 and set its style
range = para->AppendText(L"is a professional Word .NET library");
range->GetCharacterFormat()->SetFontName(L"Calibri");
range->GetCharacterFormat()->SetFontSize(15.0f);
```

---

# spire.doc cpp paragraph formatting
## demonstrates various paragraph formatting options in a Word document
```cpp
//Initialize a document
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> sec = document->AddSection();
intrusive_ptr<Paragraph> para = sec->AddParagraph();
para->AppendText(L"Paragraph Formatting");
para->ApplyStyle(BuiltinStyle::Title);

para = sec->AddParagraph();
para->AppendText(L"This paragraph is surrounded with borders.");
para->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
para->GetFormat()->GetBorders()->SetColor(Color::GetRed());

para = sec->AddParagraph();
para->AppendText(L"The alignment of this paragraph is Left.");
para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

para = sec->AddParagraph();
para->AppendText(L"The alignment of this paragraph is Center.");
para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

para = sec->AddParagraph();
para->AppendText(L"The alignment of this paragraph is Right.");
para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

para = sec->AddParagraph();
para->AppendText(L"The alignment of this paragraph is justified.");
para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Justify);

para = sec->AddParagraph();
para->AppendText(L"The alignment of this paragraph is distributed.");
para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Distribute);

para = sec->AddParagraph();
para->AppendText(L"This paragraph has the gray shadow.");
para->GetFormat()->SetBackColor(Color::GetGray());

para = sec->AddParagraph();
para->AppendText(L"This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
para->GetFormat()->SetLeftIndent(10);
para->GetFormat()->SetRightIndent(10);
para->GetFormat()->SetFirstLineIndent(15);

para = sec->AddParagraph();
para->AppendText(L"The hanging indentation of this paragraph is 15pt.");
//Negative value represents hanging indentation
para->GetFormat()->SetFirstLineIndent(-15);

para = sec->AddParagraph();
para->AppendText(L"This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
para->GetFormat()->SetAfterSpacing(20);
para->GetFormat()->SetBeforeSpacing(10);
para->GetFormat()->SetLineSpacingRule(LineSpacingRule::AtLeast);
para->GetFormat()->SetLineSpacing(10);
```

---

# spire.doc cpp list style
## restart list numbering in word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Create a new section
intrusive_ptr<Section> section = document->AddSection();

//Create first numbered list
intrusive_ptr<ListStyle> numberList = new ListStyle(document, ListType::Numbered);
numberList->SetName(L"Numbered1");
document->GetListStyles()->Add(numberList);

//Add paragraph and apply the list style
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
paragraph->AppendText(L"List Item 1");
paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

//Create second numbered list with restart numbering
intrusive_ptr<ListStyle> numberList2 = new ListStyle(document, ListType::Numbered);
numberList2->SetName(L"Numbered2");
//set start number of second list
numberList2->GetLevels()->GetItem(0)->SetStartAt(10);
document->GetListStyles()->Add(numberList2);

//Add paragraph and apply the list style
paragraph = section->AddParagraph();
paragraph->AppendText(L"List Item 5");
paragraph->GetListFormat()->ApplyStyle(numberList2->GetName());
```

---

# spire.doc cpp style retrieval
## retrieve style names from document paragraphs
```cpp
//Traverse all paragraphs in the document and get their style names through StyleName property
std::wstring styleName = L"";
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
        styleName.append(paragraph->GetStyleName());
        styleName.append(L"\n");
    }
}
```

---

# spire.doc cpp styles
## create and apply document styles
```cpp
//Initialize a document
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> sec = document->AddSection();

//Add default title style to document and modify
intrusive_ptr<Style> titleStyle = document->AddStyle(BuiltinStyle::Title);

titleStyle->GetCharacterFormat()->SetFontName(L"cambria");
titleStyle->GetCharacterFormat()->SetFontSize(28);
titleStyle->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));

//Judge if it is Paragraph Style and then set paragraph format
if (Object::Dynamic_cast<ParagraphStyle>(titleStyle) != nullptr)
{
    intrusive_ptr<ParagraphStyle> ps = Object::Dynamic_cast<ParagraphStyle>(titleStyle);
    ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
    ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetColor(Color::FromArgb(42, 123, 136));
    ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetLineWidth(1.5f);
    ps->GetParagraphFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
}

//Add default normal style and modify
intrusive_ptr<Style> normalStyle = document->AddStyle(BuiltinStyle::Normal);

normalStyle->GetCharacterFormat()->SetFontName(L"cambria");
normalStyle->GetCharacterFormat()->SetFontSize(11);

//Add default heading1 style
intrusive_ptr<Style> heading1Style = document->AddStyle(BuiltinStyle::Heading1);

heading1Style->GetCharacterFormat()->SetFontName(L"cambria");
heading1Style->GetCharacterFormat()->SetFontSize(14);
heading1Style->GetCharacterFormat()->SetBold(true);
heading1Style->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));

//Add default heading2 style
intrusive_ptr<Style> heading2Style = document->AddStyle(BuiltinStyle::Heading2);

heading2Style->GetCharacterFormat()->SetFontName(L"cambria");
heading2Style->GetCharacterFormat()->SetFontSize(12);
heading2Style->GetCharacterFormat()->SetBold(true);

//List style
intrusive_ptr<ListStyle> bulletList = new ListStyle(document, ListType::Bulleted);

bulletList->GetCharacterFormat()->SetFontName(L"cambria");
bulletList->GetCharacterFormat()->SetFontSize(12);
bulletList->SetName(L"bulletList");
document->GetListStyles()->Add(bulletList);

//Apply the style
intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
paragraph->AppendText(L"Your Name");
paragraph->ApplyStyle(BuiltinStyle::Title);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Address, City, ST ZIP Code | Telephone | Email");
paragraph->ApplyStyle(BuiltinStyle::Normal);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Objective");
paragraph->ApplyStyle(BuiltinStyle::Heading1);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Education");
paragraph->ApplyStyle(BuiltinStyle::Heading1);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"DEGREE | DATE EARNED | SCHOOL");
paragraph->ApplyStyle(BuiltinStyle::Heading2);

paragraph = sec->AddParagraph();
paragraph->AppendText(L"Major:Text");
paragraph->GetListFormat()->ApplyStyle(L"bulletList");
```

---

# spire.doc cpp mail merge
## change locale for mail merge
```cpp
intrusive_ptr<Document> document = new Document();

// Store the current culture so it can be set back once mail merge is complete.
std::locale originalLocale = std::locale::global(std::locale(""));

std::locale germanLocale("de_DE.utf8");
std::locale::global(germanLocale);

// Format time according to the current locale
std::time_t currentTime = std::time(nullptr);
std::tm* localTime = std::localtime(&currentTime);

std::wstring timeStr;
timeStr.resize(100);
std::wcsftime(&timeStr[0], timeStr.size(), L"%c", localTime);

// Execute mail merge with field names and values
std::vector<LPCWSTR_S> fieldNames = { L"Contact Name", L"Fax", L"Date" };
std::vector<LPCWSTR_S> fieldValues = { L"John Smith", L"+1 (69) 123456", timeStr.c_str() };
document->GetMailMerge()->Execute(fieldNames, fieldValues);

std::locale::global(originalLocale);
```

---

# Spire.Doc Conditional Fields
## Execute conditional fields using mail merge
```cpp
// Create conditional IF field 1
void CreateIFField1(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"1\" ");
	paragraph->AppendText(L"\"Greater than one\" ");
	paragraph->AppendText(L"\"Less than one\"");

	intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);

	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}

// Create conditional IF field 2
void CreateIFField2(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	paragraph->AppendField(L"Age", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"50\" ");
	paragraph->AppendText(L"\"The old man\" ");
	paragraph->AppendText(L"\"The young man\"");

	intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);

	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}

// Main execution of conditional fields
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> section = doc->AddSection();
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

CreateIFField1(doc, paragraph);
paragraph = section->AddParagraph();
CreateIFField2(doc, paragraph);

std::vector<LPCWSTR_S> fieldName = { L"Count", L"Age" };
std::vector<LPCWSTR_S> fieldValue = { L"2", L"30" };

doc->GetMailMerge()->Execute(fieldName, fieldValue);
doc->SetIsUpdateFields(true);
```

---

# spire.doc cpp mail merge
## hide empty regions during mail merge
```cpp
//Set the value to remove paragraphs which contain empty field.
document->GetMailMerge()->SetHideEmptyParagraphs(true);
//Set the value to remove group which contain empty field.
document->GetMailMerge()->SetHideEmptyGroup(true);
```

---

# spire.doc cpp mail merge
## identify merge field names in a Word document
```cpp
//Get the collection of group names.
std::vector<LPCWSTR_S> GroupNames = document->GetMailMerge()->GetMergeGroupNames();

//Get the collection of merge field names in a specific group.
std::vector<LPCWSTR_S> MergeFieldNamesWithinRegion = document->GetMailMerge()->GetMergeFieldNames(L"Products");

//Get the collection of all the merge field names.
std::vector<LPCWSTR_S> MergeFieldNames = document->GetMailMerge()->GetMergeFieldNames();
```

---

# Spire.Doc C++ Mail Merge
## Execute mail merge operation with field names and values
```cpp
// Create word document
intrusive_ptr<Document> document = new Document();

// Define field names for mail merge
std::vector<LPCWSTR_S> filedNames = { L"Contact Name", L"Fax", L"Date" };

// Define field values for mail merge
// C# TO C++ CONVERTER TODO TASK: There is no C++ equivalent to 'ToString':
std::vector<LPCWSTR_S> filedValues = { L"John Smith", L"+1 (69) 123456", DateTime::GetNow()->GetDate()->ToString() };

// Execute mail merge operation
document->GetMailMerge()->Execute(filedNames, filedValues);
```

---

# spire.doc cpp mail merge
## execute mail merge with field switches
```cpp
intrusive_ptr<Document> doc = new Document();

std::vector<LPCWSTR_S> fieldName = {L"XX_Name"};
std::vector<LPCWSTR_S> fieldValue = {L"Jason Tang"};

doc->GetMailMerge()->Execute(fieldName, fieldValue);
```

---

# spire.doc cpp bookmark
## copy bookmark content in word document
```cpp
//Get the bookmark by name.
intrusive_ptr<Bookmark> bookmark = doc->GetBookmarks()->GetItem(L"Test");
intrusive_ptr<DocumentObject> docObj = nullptr;

//Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
//Then need to find its outermost parent object(Table),
//and get the start/end index of current object on GetBody().
if ((Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner()))->GetIsInCell())
{
    docObj = bookmark->GetBookmarkStart()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
}
else
{
    docObj = bookmark->GetBookmarkStart()->GetOwner();
}
int startIndex = doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(docObj);
if ((Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner()))->GetIsInCell())
{
    docObj = bookmark->GetBookmarkEnd()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
}
else
{
    docObj = bookmark->GetBookmarkEnd()->GetOwner();
}
int endIndex = doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(docObj);

//Get the start/end index of the bookmark object on the paragraph.
intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner());
int pStartIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner());
int pEndIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

//Get the content of current bookmark and copy.
intrusive_ptr<TextBodySelection> select = new TextBodySelection(doc->GetSections()->GetItemInSectionCollection(0)->GetBody(), startIndex, endIndex, pStartIndex, pEndIndex);
intrusive_ptr<TextBodyPart> body = new TextBodyPart(select);
for (int i = 0; i < body->GetBodyItems()->GetCount(); i++)
{
    doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add((body->GetBodyItems())->GetItem(i)->Clone());
}
```

---

# spire.doc cpp bookmark
## create simple and nested bookmarks in Word document
```cpp
void CreateBookmark(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	
	// Create simple bookmark
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"SimpleCreateBookmark");
	paragraph->AppendText(L"This is a simple bookmark.");
	paragraph->AppendBookmarkEnd(L"SimpleCreateBookmark");

	// Create nested bookmarks
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"Root");
	paragraph->AppendText(L" This is Root data ");
	paragraph->AppendBookmarkStart(L"NestedLevel1");
	paragraph->AppendText(L" This is Nested Level1 ");
	paragraph->AppendBookmarkStart(L"NestedLevel2");
	paragraph->AppendText(L" This is Nested Level2 ");
	paragraph->AppendBookmarkEnd(L"NestedLevel2");
	paragraph->AppendBookmarkEnd(L"NestedLevel1");
	paragraph->AppendBookmarkEnd(L"Root");
}
```

---

# spire.doc cpp bookmark
## create bookmark for table in word document
```cpp
void CreateBookmarkForTable(intrusive_ptr<Document> doc, intrusive_ptr<Section> section)
{
	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append text for added paragraph
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example demonstrates how to create bookmark for a table in a Word document.");

	//Set the font in italic
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Append bookmark start
	paragraph->AppendBookmarkStart(L"CreateBookmark");

	//Append bookmark end
	paragraph->AppendBookmarkEnd(L"CreateBookmark");

	//Add table
	intrusive_ptr<Table> table = section->AddTable(true);

	//Set the number of rows and columns
	table->ResetCells(2, 2);

	//Append text for table cells		
	intrusive_ptr<TextRange> range = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"sampleA");
	range = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"sampleB");
	range = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"120");
	range = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"260");

	//Get the bookmark by index.
	intrusive_ptr<Bookmark> bookmark = doc->GetBookmarks()->GetItem(0);

	//Locate the bookmark by name.
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
	navigator->MoveToBookmark(bookmark->GetName());

	//Add table to TextBodyPart
	intrusive_ptr<TextBodyPart> part = navigator->GetBookmarkContent();
	part->GetBodyItems()->Add(table);

	//Replace bookmark cotent with table
	navigator->ReplaceBookmarkContent(part);
}
```

---

# c++ extract bookmark text
## Extract text content from a bookmark in a document
```cpp
//Creates a BookmarkNavigator instance to access the bookmark
intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
//Locate a specific bookmark by bookmark name
navigator->MoveToBookmark(L"Content");
intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

//Iterate through the items in the bookmark content to get the text
std::wstring text = L"";
for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
{
    intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
    if (Object::CheckType<Paragraph>(item))
    {
        intrusive_ptr<Paragraph> paragraph = boost::dynamic_pointer_cast<Paragraph>(item);
        for (int j = 0; j < paragraph->GetChildObjects()->GetCount(); j++)
        {
            intrusive_ptr<DocumentObject> childObject = paragraph->GetChildObjects()->GetItem(j);
            if (Object::CheckType<TextRange>(childObject))
            {
                text += (boost::dynamic_pointer_cast<TextRange>(childObject))->GetText();
            }
        }
    }
}
```

---

# spire.doc cpp bookmarks
## Get bookmarks from a Word document
```cpp
using namespace Spire::Doc;

//Create word document
intrusive_ptr<Document> document = new Document();

//Get the bookmark by index.
intrusive_ptr<Bookmark> bookmark1 = document->GetBookmarks()->GetItem(0);

//Get the bookmark by name.
intrusive_ptr<Bookmark> bookmark2 = document->GetBookmarks()->GetItem(L"Test2");
```

---

# spire.doc cpp bookmark
## insert document content at bookmark location
```cpp
//Create the first document
intrusive_ptr<Document> document1 = new Document();

//Create the second document
intrusive_ptr<Document> document2 = new Document();

//Get the first section of the first document 
intrusive_ptr<Section> section1 = document1->GetSections()->GetItemInSectionCollection(0);

//Locate the bookmark
intrusive_ptr<BookmarksNavigator> bn = new BookmarksNavigator(document1);

//Find bookmark by name
bn->MoveToBookmark(L"Test", true, true);

//Get bookmarkStart
intrusive_ptr<BookmarkStart> start = bn->GetCurrentBookmark()->GetBookmarkStart();

//Get the owner paragraph
intrusive_ptr<Paragraph> para = start->GetOwnerParagraph();

//Get the para index
int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

//Insert the paragraphs of document2
for (int i = 0; i < document2->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section2 = document2->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> paragraph = section2->GetParagraphs()->GetItemInParagraphCollection(j);
        section1->GetBody()->GetChildObjects()->Insert(index + 1, Object::Dynamic_cast<Paragraph>(paragraph->Clone()));
    }
}
```

---

# spire.doc cpp bookmark
## insert image at bookmark location
```cpp
//Create a document instance
intrusive_ptr<Document> doc = new Document();

//Create an instance of BookmarksNavigator
intrusive_ptr<BookmarksNavigator> bn = new BookmarksNavigator(doc);

//Find a bookmark named Test
bn->MoveToBookmark(L"Test", true, true);

//Add a section
intrusive_ptr<Section> section0 = doc->AddSection();

//Add a paragraph for the section
intrusive_ptr<Paragraph> paragraph = section0->AddParagraph();

//Add a picture into the paragraph
intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(L"Word.png");
//Add the paragraph at the position of bookmark
bn->InsertParagraph(paragraph);

//Remove the section0
doc->GetSections()->Remove(section0);
```

---

# spire.doc cpp bookmark
## remove content from bookmark
```cpp
//Get the bookmark by name.            
intrusive_ptr<Bookmark> bookmark = document->GetBookmarks()->GetItem(L"Test");

intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner());
int startIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner());
int endIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

//Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
//This method is only to remove the content of the bookmark.
for (int i = startIndex + 1; i < endIndex; i++)
{
    para->GetChildObjects()->RemoveAt(startIndex + 1);
}
```

---

# Spire.Doc C++ Bookmark
## Remove bookmark from document
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();

//Get the bookmark by name
intrusive_ptr<Bookmark> bookmark = document->GetBookmarks()->GetItem(L"Test");

//Remove the bookmark, not its content
document->GetBookmarks()->Remove(bookmark);
```

---

# spire.doc cpp bookmark
## replace bookmark content in document
```cpp
using namespace Spire::Doc;

//Create document object
intrusive_ptr<Document> doc = new Document();

//Create bookmark navigator
intrusive_ptr<BookmarksNavigator> bookmarkNavigator = new BookmarksNavigator(doc);

//Locate the bookmark
bookmarkNavigator->MoveToBookmark(L"Test");

//Replace the bookmark content with new content
bookmarkNavigator->ReplaceBookmarkContent(L"This is replaced content.", false);

//Close the document
doc->Close();
```

---

# spire.doc cpp bookmark
## replace bookmark content with table
```cpp
//Create a table
intrusive_ptr<Table> table = new Table(doc, true);

//Create data
const int rowsCount = 4;
const int colsCount = 5;
std::wstring data[rowsCount][colsCount] = {
	{L"Name", L"Capital", L"Continent", L"Area", L"Population"},
	{L"Argentina", L"Buenos Aires", L"South America", L"2777815", L"32300003"},
	{L"Bolivia", L"La Paz", L"South America", L"1098575", L"7300000"},
	{L"Brazil", L"Brasilia", L"South America", L"8511196", L"150400000"}
};
table->ResetCells(rowsCount, colsCount);

//Fill the table with the data
for (int i = 0; i < rowsCount; i++)
{
	for (int j = 0; j < colsCount; j++)
	{
		table->GetRows()->GetItemInRowCollection(i)->GetCells()->GetItemInCellCollection(j)->AddParagraph()->AppendText(data[i][j].c_str());
	}
}

//Get the specific bookmark by its name
intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
navigator->MoveToBookmark(L"Test");

//Create a TextBodyPart instance and add the table to it
intrusive_ptr<TextBodyPart> part = new TextBodyPart(doc);
part->GetBodyItems()->Add(table);

//Replace the current bookmark content with the TextBodyPart object
navigator->ReplaceBookmarkContent(part);
```

---

# spire.doc cpp comment
## add comment for specific text in document
```cpp
void InsertComments(intrusive_ptr<Document> doc, const std::wstring& keystring)
{
	//Find the key string
	intrusive_ptr<TextSelection> find = doc->FindString(keystring.c_str(), false, true);

	//Create the commentmarkStart and commentmarkEnd
	intrusive_ptr<CommentMark> commentmarkStart = new CommentMark(doc);
	commentmarkStart->SetType(CommentMarkType::CommentStart);
	intrusive_ptr<CommentMark> commentmarkEnd = new CommentMark(doc);
	commentmarkEnd->SetType(CommentMarkType::CommentEnd);

	//Add the content for comment
	intrusive_ptr<Comment> comment = new Comment(doc);
	comment->GetBody()->AddParagraph()->SetText(L"Test comments");
	comment->GetFormat()->SetAuthor(L"E-iceblue");

	//Get the textRange
	intrusive_ptr<TextRange> range = find->GetAsOneRange();

	//Get its paragraph
	intrusive_ptr<Paragraph> para = range->GetOwnerParagraph();

	//Get the index of textRange 
	int index = para->GetChildObjects()->IndexOf(range);

	//Add comment
	para->GetChildObjects()->Add(comment);

	//Insert the commentmarkStart and commentmarkEnd
	para->GetChildObjects()->Insert(index, commentmarkStart);
	para->GetChildObjects()->Insert(index + 2, commentmarkEnd);
}
```

---

# spire.doc cpp comment
## insert comments into document
```cpp
void InsertComments(intrusive_ptr<Section> section)
{
	//Insert comment.
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(1);
	intrusive_ptr<Spire::Doc::Comment> comment = paragraph->AppendComment(L"Spire.Doc for .NET");
	comment->GetFormat()->SetAuthor(L"E-iceblue");
	comment->GetFormat()->SetInitial(L"CM");
}
```

---

# Spire.Doc C++ Comment Extraction
## Extract text content from all comments in a document
```cpp
std::wstring stringBuilder;

//Traverse all comments
for (int i = 0; i < doc->GetComments()->GetCount(); i++)
{
	intrusive_ptr<Comment> comment = doc->GetComments()->GetItem(i);
	for (int j = 0; j < comment->GetBody()->GetParagraphs()->GetCount(); j++)
	{
		intrusive_ptr<Paragraph> p = comment->GetBody()->GetParagraphs()->GetItemInParagraphCollection(j);
		stringBuilder.append(p->GetText());
		stringBuilder.append(L"\n");
	}
}
```

---

# Spire.Doc C++ Comment Picture
## Insert a picture into a document comment
```cpp
//Get the first paragraph and insert comment
intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(2);
intrusive_ptr<Comment> comment = paragraph->AppendComment(L"This is a comment.");
comment->GetFormat()->SetAuthor(L"E-iceblue");

//Load a picture
intrusive_ptr<DocPicture> docPicture = new DocPicture(doc);
docPicture->LoadImageSpire(DATAPATH"/E-iceblue.png");

//Insert the picture into the comment GetBody()
comment->GetBody()->AddParagraph()->GetChildObjects()->Add(docPicture);
```

---

# Spire.Doc C++ Comment Operations
## Remove and replace comments in a Word document
```cpp
//Create a document object
intrusive_ptr<Document> doc = new Document();

//Replace the content of the first comment
doc->GetComments()->GetItem(0)->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0)->Replace(L"This is the title", L"This comment is changed.", false, false);

//Remove the second comment
doc->GetComments()->RemoveAt(1);
```

---

# spire.doc cpp comment
## remove content with comment
```cpp
//Get the first comment
intrusive_ptr<Comment> comment = document->GetComments()->GetItem(0);

//Get the paragraph of obtained comment
intrusive_ptr<Paragraph> para = comment->GetOwnerParagraph();

//Get index of the CommentMarkStart 
int startIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkStart());

//Get index of the CommentMarkEnd
int endIndex = para->GetChildObjects()->IndexOf(comment->GetCommentEnd());

//Create a list
std::vector<intrusive_ptr<TextRange>> list;

//Get TextRanges between the indexes
for (int i = startIndex; i < endIndex; i++)
{
    if (Object::CheckType<TextRange>(para->GetChildObjects()->GetItem(i)))
    {
        list.push_back(boost::dynamic_pointer_cast<TextRange>(para->GetChildObjects()->GetItem(i)));
    }
}

//Insert a new TextRange
intrusive_ptr<TextRange> textRange = new TextRange(document);

//Set text is null
textRange->SetText(nullptr);

//Insert the new textRange
para->GetChildObjects()->Insert(endIndex, textRange);

//Remove previous TextRanges
for (size_t i = 0; i < list.size(); i++)
{
    para->GetChildObjects()->Remove(list[i]);
}
```

---

# spire.doc cpp comment
## reply to comment and add picture
```cpp
//get the first comment.
intrusive_ptr<Comment> comment1 = doc->GetComments()->GetItem(0);

//create a new comment and specify the author and content.
intrusive_ptr<Comment> replyComment1 = new Comment(doc);
replyComment1->GetFormat()->SetAuthor(L"E-iceblue");
replyComment1->GetBody()->AddParagraph()->AppendText(L"Spire.Doc is a professional Word  library on operating Word documents.");

//add the new comment as a reply to the selected comment.
comment1->ReplyToComment(replyComment1);

intrusive_ptr<DocPicture> docPicture = new DocPicture(doc);

//insert a picture in the comment
replyComment1->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Add(docPicture);
```

---

# spire.doc cpp barcode
## add barcode image to word document
```cpp
//Add barcode image
intrusive_ptr<DocPicture> picture = document->GetSections()->GetItemInSectionCollection(0)->AddParagraph()->AppendPicture(imgPath.c_str());
```

---

# spire.doc cpp horizontal line
## add horizontal line to Word document
```cpp
//Create Word document.
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> sec = doc->AddSection();
intrusive_ptr<Paragraph> para = sec->AddParagraph();
para->AppendHorizonalLine();
```

---

# spire.doc cpp image manipulation
## add image and textbox to document footer
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Add a picture in footer and set it's position
intrusive_ptr<DocPicture> picture = document->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendPicture(imgPath.c_str());
picture->SetVerticalOrigin(VerticalOrigin::Page);
picture->SetHorizontalOrigin(HorizontalOrigin::Page);
picture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);
picture->SetTextWrappingStyle(TextWrappingStyle::None);

//Add a textbox in footer and set it's position
intrusive_ptr<TextBox> textbox = document->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendTextBox(150, 20);
textbox->SetVerticalOrigin(VerticalOrigin::Page);
textbox->SetHorizontalOrigin(HorizontalOrigin::Page);
textbox->SetHorizontalPosition(300);
textbox->SetVerticalPosition(700);
textbox->GetBody()->AddParagraph()->AppendText(L"Welcome to E-iceblue");
```

---

# spire.doc cpp shape group
## create a shape group with multiple shapes including text boxes and arrows
```cpp
//create a document
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> sec = doc->AddSection();

//add a new paragraph
intrusive_ptr<Paragraph> para = sec->AddParagraph();
//add a shape group with the height and width
intrusive_ptr<ShapeGroup> shapegroup = para->AppendShapeGroup(375, 462);
shapegroup->SetHorizontalPosition(180);
//calculate the scale ratio
float X = static_cast<float>(shapegroup->GetWidth() / 1000.0f);
float Y = static_cast<float>(shapegroup->GetHeight() / 1000.0f);

intrusive_ptr<TextBox> txtBox = new TextBox(doc);
txtBox->SetShapeType(ShapeType::RoundRectangle);
txtBox->SetWidth(125 / X);
txtBox->SetHeight(54 / Y);
intrusive_ptr<Paragraph> paragraph = txtBox->GetBody()->AddParagraph();
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
paragraph->AppendText(L"Start");
txtBox->SetHorizontalPosition(19 / X);
txtBox->SetVerticalPosition(27 / Y);
txtBox->GetFormat()->SetLineColor(Color::GetGreen());
shapegroup->GetChildObjects()->Add(txtBox);

intrusive_ptr<ShapeObject> arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
arrowLineShape->SetWidth(16 / X);
arrowLineShape->SetHeight(40 / Y);
arrowLineShape->SetHorizontalPosition(69 / X);
arrowLineShape->SetVerticalPosition(87 / Y);
arrowLineShape->SetStrokeColor(Color::GetPurple());
shapegroup->GetChildObjects()->Add(arrowLineShape);

txtBox = new TextBox(doc);
txtBox->SetShapeType(ShapeType::Rectangle);
txtBox->SetWidth(125 / X);
txtBox->SetHeight(54 / Y);
paragraph = txtBox->GetBody()->AddParagraph();
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
paragraph->AppendText(L"Step 1");
txtBox->SetHorizontalPosition(19 / X);
txtBox->SetVerticalPosition(131 / Y);
txtBox->GetFormat()->SetLineColor(Color::GetBlue());
shapegroup->GetChildObjects()->Add(txtBox);

arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
arrowLineShape->SetWidth(16 / X);
arrowLineShape->SetHeight(40 / Y);
arrowLineShape->SetHorizontalPosition(69 / X);
arrowLineShape->SetVerticalPosition(192 / Y);
arrowLineShape->SetStrokeColor(Color::GetPurple());
shapegroup->GetChildObjects()->Add(arrowLineShape);

txtBox = new TextBox(doc);
txtBox->SetShapeType(ShapeType::Parallelogram);
txtBox->SetWidth(149 / X);
txtBox->SetHeight(59 / Y);
paragraph = txtBox->GetBody()->AddParagraph();
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
paragraph->AppendText(L"Step 2");
txtBox->SetHorizontalPosition(7 / X);
txtBox->SetVerticalPosition(236 / Y);
txtBox->GetFormat()->SetLineColor(Color::GetBlueViolet());
shapegroup->GetChildObjects()->Add(txtBox);

arrowLineShape = new ShapeObject(doc, ShapeType::DownArrow);
arrowLineShape->SetWidth(16 / X);
arrowLineShape->SetHeight(40 / Y);
arrowLineShape->SetHorizontalPosition(66 / X);
arrowLineShape->SetVerticalPosition(300 / Y);
arrowLineShape->SetStrokeColor(Color::GetPurple());
shapegroup->GetChildObjects()->Add(arrowLineShape);

txtBox = new TextBox(doc);
txtBox->SetShapeType(ShapeType::Rectangle);
txtBox->SetWidth(125 / X);
txtBox->SetHeight(54 / Y);
paragraph = txtBox->GetBody()->AddParagraph();
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
paragraph->AppendText(L"Step 3");
txtBox->SetHorizontalPosition(19 / X);
txtBox->SetVerticalPosition(345 / Y);
txtBox->GetFormat()->SetLineColor(Color::GetBlue());
shapegroup->GetChildObjects()->Add(txtBox);
```

---

# spire.doc cpp shapes
## add various shapes to word document
```cpp
//Create Word document.
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> sec = doc->AddSection();
intrusive_ptr<Paragraph> para = sec->AddParagraph();
int x = 60, y = 40, lineCount = 0;
for (int i = 1; i < 20; i++)
{
    if (lineCount > 0 && lineCount % 8 == 0)
    {
        para->AppendBreak(BreakType::PageBreak);
        x = 60;
        y = 40;
        lineCount = 0;
    }
    //Add shape and set its size and position.
    intrusive_ptr<ShapeObject> shape = para->AppendShape(50, 50, (ShapeType)i);
    shape->SetHorizontalOrigin(HorizontalOrigin::Page);
    shape->SetHorizontalPosition(x);
    shape->SetVerticalOrigin(VerticalOrigin::Page);
    shape->SetVerticalPosition(y + 50);
    x = x + static_cast<int>(shape->GetWidth()) + 50;
    if (i > 0 && i % 5 == 0)
    {
        y = y + static_cast<int>(shape->GetHeight()) + 120;
        lineCount++;
        x = 60;
    }
}
```

---

# spire.doc cpp shape alignment
## align shapes to center horizontally in document
```cpp
intrusive_ptr<Document> doc = new Document();

intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
{
	intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
	for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
	{
		intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
		if (Object::CheckType<ShapeObject>(obj))
		{
			//Set the horizontal alignment as center
			(boost::dynamic_pointer_cast<ShapeObject>(obj))->SetHorizontalAlignment(ShapeHorizontalAlignment::Center);
		}
	}
}
```

---

# spire.doc cpp image extraction
## extract images from document
```cpp
//open document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile("document_path");

//document elements, each of them has child elements
std::deque<intrusive_ptr<ICompositeObject>> nodes;
nodes.push_back(document);

//embedded images list.
std::vector<std::vector<byte>> images;
//traverse
while (nodes.size() > 0)
{
	intrusive_ptr<ICompositeObject> node = nodes.front();
	nodes.pop_front();
	for (int i = 0; i < node->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<IDocumentObject> child = node->GetChildObjects()->GetItem(i);
		if (child->GetDocumentObjectType() == DocumentObjectType::Picture)
		{
			intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(child);
			std::vector<byte> imageByte = picture->GetImageBytes();
			images.push_back(imageByte);
		}
		else if (Object::CheckType<ICompositeObject>(child))
		{
			nodes.push_back(boost::dynamic_pointer_cast<ICompositeObject>(child));
		}
	}
}
document->Close();
```

---

# spire.doc cpp alternative text
## Extract alternative text from shapes in a document
```cpp
//Loop through shapes and get the AlternativeText
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(j);
        for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
        {
            intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
            if (Object::CheckType<ShapeObject>(obj))
            {
                std::wstring text = (boost::dynamic_pointer_cast<ShapeObject>(obj))->GetAlternativeText();
                //Append the alternative text in builder
                builder.append(text);
                builder.append(L"\n");
            }
        }
    }
}
```

---

# spire.doc cpp image
## Insert image into Word document
```cpp
void InsertImage(intrusive_ptr<Section> section)
{
	//Add paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	//Add a image and set its width and height
	intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(DATAPATH"/Spire.Doc.png");

	picture->SetWidth(100);
	picture->SetHeight(100);
}

//Create a document
intrusive_ptr<Document> document = new Document();

//Add a seciton
intrusive_ptr<Section> section = document->AddSection();

//insert image
InsertImage(section);
```

---

# spire.doc cpp image insertion
## insert an image into a word document
```cpp
//Create a picture
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
picture->LoadImageSpire(DATAPATH"/Word.png");
//set image's position
picture->SetHorizontalPosition(50.0F);
picture->SetVerticalPosition(60.0F);

//set image's size
picture->SetWidth(200);
picture->SetHeight(200);

//set textWrappingStyle with image
picture->SetTextWrappingStyle(TextWrappingStyle::Through);
//Insert the picture at the beginning of the second paragraph
paragraph->GetChildObjects()->Insert(0, picture);
```

---

# spire.doc cpp wordart
## insert WordArt into document
```cpp
//Add a paragraph.
intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->AddParagraph();

//Add a shape.
intrusive_ptr<ShapeObject> shape = paragraph->AppendShape(250, 70, ShapeType::TextWave4);

//Set the position of the shape.
shape->SetVerticalPosition(20);
shape->SetHorizontalPosition(80);

//Set the text of WordArt.
shape->GetWordArt()->SetText(L"Thanks for reading.");

//Set the fill color.
shape->SetFillColor(Color::GetRed());

//Set the border color of the text.
shape->SetStrokeColor(Color::GetYellow());
```

---

# spire.doc cpp shape removal
## remove shapes from word document
```cpp
using namespace Spire::Doc;

intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get all the child objects of paragraph
for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
{
    intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
    for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
    {
        //If the child objects is shape object
        if (Object::CheckType<ShapeObject>(para->GetChildObjects()->GetItem(j)))
        {
            //Remove the shape object
            para->GetChildObjects()->RemoveAt(j);
            --j;
        }
    }
}
```

---

# spire.doc cpp image replacement
## Replace images with text in a Word document
```cpp
//Replace all pictures with texts
int j = 1;
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
    for (int k = 0; k < sec->GetParagraphs()->GetCount(); k++)
    {
        intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(k);
        std::vector<intrusive_ptr<DocumentObject>> pictures;
        //Get all pictures in the Word document
        for (int m = 0; m < para->GetChildObjects()->GetCount(); m++)
        {
            intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(m);
            if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
            {
                pictures.push_back(docObj);
            }
        }

        //Replace pictures with the text "Here was image {image index}"
        for (auto pic : pictures)
        {
            int index = para->GetChildObjects()->IndexOf(pic);
            intrusive_ptr<TextRange> range = new TextRange(doc);
            wstring temp = L"Here was image " + to_wstring(j) + L"";
            range->SetText(temp.c_str());
            para->GetChildObjects()->Insert(index, range);
            para->GetChildObjects()->Remove(pic);
            j++;
        }
    }
}
```

---

# spire.doc cpp image
## reset image size in document
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
//Get the first paragraph
intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(0);

//Reset the image size of the first paragraph
for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> docObj = paragraph->GetChildObjects()->GetItem(i);
	if (Object::CheckType<DocPicture>(docObj))
	{
		intrusive_ptr<DocPicture> picture = boost::dynamic_pointer_cast<DocPicture>(docObj);
		picture->SetWidth(50.0f);
		picture->SetHeight(50.0f);
	}
}
```

---

# spire.doc cpp shape
## reset shape size in document
```cpp
using namespace Spire::Doc;

//Get the first section and the first paragraph that contains the shape
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(0);

//Get the second shape and reset the width and height for the shape
intrusive_ptr<ShapeObject> shape = Object::Dynamic_cast<ShapeObject>(para->GetChildObjects()->GetItem(1));
shape->SetWidth(200);
shape->SetHeight(200);
```

---

# spire.doc cpp shape rotation
## rotate shape objects in word document
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Traverse the word document and set the shape rotation as 20
for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
{
    intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
    for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
        if (Object::CheckType<ShapeObject>(obj))
        {
            (boost::dynamic_pointer_cast<ShapeObject>(obj))->SetRotation(20.0);
        }
    }
}
```

---

# spire.doc cpp image text wrap
## set text wrapping style for images in word document
```cpp
//Load Document
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());

for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
	intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
	for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
	{
		intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(j);
		std::vector<intrusive_ptr<DocumentObject>> pictures;
		//Get all pictures in the Word document
		for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
		{
			intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(k);
			if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
			{
				pictures.push_back(docObj);
			}
		}

		//Set text wrap styles for each picture
		for (auto pic : pictures)
		{
			intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(pic);
			picture->SetTextWrappingStyle(TextWrappingStyle::Through);
			picture->SetTextWrappingType(TextWrappingType::Both);
		}
	}
}

//Save and launch document
doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
doc->Close();
```

---

# spire.doc cpp image transparency
## set transparent color for images in document
```cpp
//Get the first paragraph in the first section
intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

//Set the blue color of the image(s) in the paragraph to transparent
for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
{
    intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(i);
    if (Object::CheckType<DocPicture>(obj))
    {
        intrusive_ptr<DocPicture> picture = boost::dynamic_pointer_cast<DocPicture>(obj);
        picture->SetTransparentColor(Color::GetBlue());
    }
}
```

---

# spire.doc cpp image update
## update image in word document
```cpp
//Get all pictures in the Word document
std::vector<intrusive_ptr<DocumentObject>> pictures;
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
	intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
	for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
	{
		intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(j);
		for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
		{
			intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(k);
			if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
			{
				pictures.push_back(docObj);
			}
		}
	}
}

//Replace the first picture with a new image file
intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(pictures[0]);
picture->LoadImageSpire(DATAPATH"/E-iceblue.png");
```

---

# spire.doc cpp header footer
## add header only to first page
```cpp
//Get the header from the first section
intrusive_ptr<HeaderFooter> header = doc1->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

//Get the first page header of the destination document
intrusive_ptr<HeaderFooter> firstPageHeader = doc2->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFirstPageHeader();

//Specify that the current section has a different header/footer for the first page
for (int i = 0; i < doc2->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc2->GetSections()->GetItemInSectionCollection(i);
    section->GetPageSetup()->SetDifferentFirstPageHeaderFooter(true);
}

//Removes all child objects in firstPageHeader
firstPageHeader->GetParagraphs()->Clear();

//Add all child objects of the header to firstPageHeader
for (int j = 0; j < header->GetChildObjects()->GetCount(); j++)
{
    intrusive_ptr<DocumentObject> obj = header->GetChildObjects()->GetItem(j);
    firstPageHeader->GetChildObjects()->Add(obj->Clone());
}
```

---

# spire.doc cpp header footer
## adjust header and footer height in document
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Adjust the height of headers in the section
section->GetPageSetup()->SetHeaderDistance(100);

//Adjust the height of footers in the section
section->GetPageSetup()->SetFooterDistance(100);
```

---

# spire.doc cpp header footer
## copy header from one document to another
```cpp
//Get the header section from the source document
intrusive_ptr<HeaderFooter> header = doc1->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

//Copy each object in the header of source file to destination file
for (int i = 0; i < doc2->GetSections()->GetCount(); i++)
{
	intrusive_ptr<Section> section = doc2->GetSections()->GetItemInSectionCollection(i);
	for (int j = 0; j < header->GetChildObjects()->GetCount(); j++)
	{
		intrusive_ptr<DocumentObject> obj = header->GetChildObjects()->GetItem(j);
		section->GetHeadersFooters()->GetHeader()->GetChildObjects()->Add(obj->Clone());
	}
}
```

---

# spire.doc cpp header footer
## set different first page header and footer
```cpp
//Get the section and set the property true
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
section->GetPageSetup()->SetDifferentFirstPageHeaderFooter(true);

//Set the first page header. Here we append a picture in the header
intrusive_ptr<Paragraph> paragraph1 = section->GetHeadersFooters()->GetFirstPageHeader()->AddParagraph();
paragraph1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

intrusive_ptr<DocPicture> headerimage = paragraph1->AppendPicture(L"E-iceblue.png");

//Set the first page footer
intrusive_ptr<Paragraph> paragraph2 = section->GetHeadersFooters()->GetFirstPageFooter()->AddParagraph();
paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
intrusive_ptr<TextRange> FF = paragraph2->AppendText(L"First Page Footer");
FF->GetCharacterFormat()->SetFontSize(10);

//Set the other header & footer. If you only need the first page header & footer, don't set this
intrusive_ptr<Paragraph> paragraph3 = section->GetHeadersFooters()->GetHeader()->AddParagraph();
paragraph3->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
intrusive_ptr<TextRange> NH = paragraph3->AppendText(L"Spire.Doc for .NET");
NH->GetCharacterFormat()->SetFontSize(10);

intrusive_ptr<Paragraph> paragraph4 = section->GetHeadersFooters()->GetFooter()->AddParagraph();
paragraph4->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
intrusive_ptr<TextRange> NF = paragraph4->AppendText(L"E-iceblue");
NF->GetCharacterFormat()->SetFontSize(10);
```

---

# spire.doc cpp header footer
## insert header and footer with images and page numbers
```cpp
void InsertHeaderAndFooter(intrusive_ptr<Section> section)
{
	intrusive_ptr<HeaderFooter> header = section->GetHeadersFooters()->GetHeader();
	intrusive_ptr<HeaderFooter> footer = section->GetHeadersFooters()->GetFooter();

	//insert picture and text to header
	intrusive_ptr<Paragraph> headerParagraph = header->AddParagraph();

	intrusive_ptr<DocPicture> headerPicture = headerParagraph->AppendPicture(DATAPATH"/Header.png");

	//header text
	intrusive_ptr<TextRange> text = headerParagraph->AppendText(L"Demo of Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(10);
	text->GetCharacterFormat()->SetItalic(true);
	headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetSpace(0.05F);

	//header picture layout - text wrapping
	headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);

	//header picture layout - position
	headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

	//insert picture to footer
	intrusive_ptr<Paragraph> footerParagraph = footer->AddParagraph();

	intrusive_ptr<DocPicture> footerPicture = footerParagraph->AppendPicture(DATAPATH"/Footer.png");

	//footer picture layout
	footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
	footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//insert page number
	footerParagraph->AppendField(L"page number", FieldType::FieldPage);
	footerParagraph->AppendText(L" of ");
	footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
	footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetSpace(0.05F);
}
```

---

# spire.doc cpp header footer
## Add images to document header and footer
```cpp
//Get the header of the first page
intrusive_ptr<HeaderFooter> header = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

//Add a paragraph for the header
intrusive_ptr<Paragraph> paragraph = header->AddParagraph();

//Set the format of the paragraph
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

//Append a picture in the paragraph
intrusive_ptr<DocPicture> headerimage = paragraph->AppendPicture(L"path/to/header/image.png");
headerimage->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

//Get the footer of the first section
intrusive_ptr<HeaderFooter> footer = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetFooter();

//Add a paragraph for the footer
intrusive_ptr<Paragraph> paragraph2 = footer->AddParagraph();

//Set the format of the paragraph
paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

//Append a picture in the paragraph
intrusive_ptr<DocPicture> footerimage = paragraph2->AppendPicture(L"path/to/footer/image.png");

//Append text in the paragraph
wstring string = L"Copyright \u00A9 2013 e-iceblue. All Rights Reserved.";
intrusive_ptr<TextRange> TR = paragraph2->AppendText(string.c_str());
TR->GetCharacterFormat()->SetFontName(L"Arial");
TR->GetCharacterFormat()->SetFontSize(10);
TR->GetCharacterFormat()->SetTextColor(Color::GetBlack());
```

---

# spire.doc cpp header footer protection
## Lock document content while keeping headers unlocked
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Protect the document and set the ProtectionType as AllowOnlyFormFields
doc->Protect(ProtectionType::AllowOnlyFormFields, L"123");

//Set the ProtectForm as false to unprotect the section
section->SetProtectForm(false);
```

---

# spire.doc c++ header footer
## create different headers and footers for odd and even pages
```cpp
//Get the section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Set the DifferentOddAndEvenPagesHeaderFooter property to true
section->GetPageSetup()->SetDifferentOddAndEvenPagesHeaderFooter(true);

//Add odd header
intrusive_ptr<Paragraph> P3 = section->GetHeadersFooters()->GetOddHeader()->AddParagraph();
intrusive_ptr<TextRange> OH = P3->AppendText(L"Odd Header");
P3->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
OH->GetCharacterFormat()->SetFontName(L"Arial");
OH->GetCharacterFormat()->SetFontSize(10);

//Add even header
intrusive_ptr<Paragraph> P4 = section->GetHeadersFooters()->GetEvenHeader()->AddParagraph();
intrusive_ptr<TextRange> EH = P4->AppendText(L"Even Header from E-iceblue Using Spire.Doc");
P4->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
EH->GetCharacterFormat()->SetFontName(L"Arial");
EH->GetCharacterFormat()->SetFontSize(10);

//Add odd footer
intrusive_ptr<Paragraph> P2 = section->GetHeadersFooters()->GetOddFooter()->AddParagraph();
intrusive_ptr<TextRange> OF = P2->AppendText(L"Odd Footer");
P2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
OF->GetCharacterFormat()->SetFontName(L"Arial");
OF->GetCharacterFormat()->SetFontSize(10);

//Add even footer
intrusive_ptr<Paragraph> P1 = section->GetHeadersFooters()->GetEvenFooter()->AddParagraph();
intrusive_ptr<TextRange> EF = P1->AppendText(L"Even Footer from E-iceblue Using Spire.Doc");
EF->GetCharacterFormat()->SetFontName(L"Arial");
EF->GetCharacterFormat()->SetFontSize(10);
P1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
```

---

# spire.doc cpp page border
## configure page border surround for header and footer
```cpp
//Create a new document
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> section = doc->AddSection();

//Add a sample page border to the document
section->GetPageSetup()->GetBorders()->SetBorderType(BorderStyle::Wave);
section->GetPageSetup()->GetBorders()->SetColor(Color::GetGreen());
section->GetPageSetup()->GetBorders()->GetLeft()->SetSpace(20);
section->GetPageSetup()->GetBorders()->GetRight()->SetSpace(20);

//Add a header and set its format
intrusive_ptr<Paragraph> paragraph1 = section->GetHeadersFooters()->GetHeader()->AddParagraph();
paragraph1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
intrusive_ptr<TextRange> headerText = paragraph1->AppendText(L"Header isn't included in page border");
headerText->GetCharacterFormat()->SetFontName(L"Calibri");
headerText->GetCharacterFormat()->SetFontSize(20);
headerText->GetCharacterFormat()->SetBold(true);

//Add a footer and set its format
intrusive_ptr<Paragraph> paragraph2 = section->GetHeadersFooters()->GetFooter()->AddParagraph();
paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
intrusive_ptr<TextRange> footerText = paragraph2->AppendText(L"Footer is included in page border");
footerText->GetCharacterFormat()->SetFontName(L"Calibri");
footerText->GetCharacterFormat()->SetFontSize(20);
footerText->GetCharacterFormat()->SetBold(true);

//Set the header not included in the page border while the footer included
section->GetPageSetup()->SetPageBorderIncludeHeader(false);
section->GetPageSetup()->SetHeaderDistance(40);
section->GetPageSetup()->SetPageBorderIncludeFooter(true);
section->GetPageSetup()->SetFooterDistance(40);
```

---

# spire.doc cpp footer
## remove footer from word document
```cpp
using namespace Spire::Doc;

//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Traverse the word document and clear all footers in different type
for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
{
	intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
	for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
	{
		intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
		//Clear footer in the first page
		intrusive_ptr<HeaderFooter> footer;
		footer = section->GetHeadersFooters()->GetFirstPageFooter();
		if (footer != nullptr)
		{
			footer->GetChildObjects()->Clear();
		}
		//Clear footer in the odd page
		footer = section->GetHeadersFooters()->GetOddFooter();
		if (footer != nullptr)
		{
			footer->GetChildObjects()->Clear();
		}
		//Clear footer in the even page
		footer = section->GetHeadersFooters()->GetEvenFooter();
		if (footer != nullptr)
		{
			footer->GetChildObjects()->Clear();
		}
	}
}
```

---

# spire.doc cpp header footer
## remove headers from document
```cpp
//Get the first section of the document
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Clear header in the first page
intrusive_ptr<HeaderFooter> header;
header = section->GetHeadersFooters()->GetFirstPageHeader();
if (header != nullptr)
{
    header->GetChildObjects()->Clear();
}

//Clear header in the odd page
header = section->GetHeadersFooters()->GetOddHeader();
if (header != nullptr)
{
    header->GetChildObjects()->Clear();
}

//Clear header in the even page
header = section->GetHeadersFooters()->GetEvenHeader();
if (header != nullptr)
{
    header->GetChildObjects()->Clear();
}
```

---

# spire.doc cpp table
## add alternative text to table
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table in the section
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Add alternative text
//Add title
table->SetTitle(L"Table 1");
//Add description
table->SetTableDescription(L"Description Text");
```

---

# spire.doc cpp table operations
## add or delete rows in a word table
```cpp
// Get the first section and the first table from that section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

// Delete the seventh row
table->GetRows()->RemoveAt(7);

// Add a row and insert it into specific position
intrusive_ptr<TableRow> row = new TableRow(document);
for (int i = 0; i < table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetCount(); i++)
{
	intrusive_ptr<TableCell> tc = row->AddCell();
	intrusive_ptr<Paragraph> paragraph = tc->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	paragraph->AppendText(L"Added");
}
table->GetRows()->Insert(2, row);

// Add a row at the end of table
table->AddRow();
```

---

# spire.doc cpp table
## add or remove column from table
```cpp
using namespace Spire::Doc;

void AddColumn(intrusive_ptr<Table> table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		intrusive_ptr<TableCell> addCell = new TableCell(table->GetDocument());
		table->GetRows()->GetItemInRowCollection(r)->GetCells()->Insert(columnIndex, addCell);
	}
}

void RemoveColumn(intrusive_ptr<Table> table, int columnIndex)
{
	for (int r = 0; r < table->GetRows()->GetCount(); r++)
	{
		table->GetRows()->GetItemInRowCollection(r)->GetCells()->RemoveAt(columnIndex);
	}
}
```

---

# spire.doc cpp table cell
## add picture to table cell
```cpp
//Get the first table from the first section of the document
intrusive_ptr<Table> table1 = Object::Dynamic_cast<Table>(doc->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

//Add a picture to the specified table cell and set picture size
intrusive_ptr<DocPicture> picture = table1->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->GetParagraphs()->GetItemInParagraphCollection(0)->AppendPicture(DATAPATH"/Spire.Doc.png");

picture->SetWidth(100);
picture->SetHeight(100);
```

---

# spire.doc cpp table formatting
## allow table rows to break across pages
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

for (int i = 0; i < table->GetRows()->GetCount(); i++)
{
    intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
    //Allow break across pages
    row->GetRowFormat()->SetIsBreakAcrossPages(true);
}
```

---

# spire.doc cpp table
## auto fit table to contents
```cpp
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Automatically fit the table to the cell content
table->AutoFit(AutoFitBehaviorType::AutoFitToContents);
```

---

# Spire.Doc C++ Table AutoFit
## Set table to fixed column widths
```cpp
//Get table from document
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));
//The table is set to a fixed size
table->AutoFit(AutoFitBehaviorType::FixedColumnWidths);
```

---

# spire.doc cpp table
## AutoFit table to window
```cpp
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));
//Automatically fit the table to the active window width
table->AutoFit(AutoFitBehaviorType::AutoFitToWindow);
```

---

# Spire.Doc C++ Table Cell Merge Status
## Check and report merge status of cells in a Word document table
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table in the section
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

std::wstring stringBuidler;
for (int i = 0; i < table->GetRows()->GetCount(); i++)
{
    intrusive_ptr<TableRow> tableRow = table->GetRows()->GetItemInRowCollection(i);
    for (int j = 0; j < tableRow->GetCells()->GetCount(); j++)
    {
        intrusive_ptr<TableCell> tableCell = tableRow->GetCells()->GetItemInCellCollection(j);
        CellMerge verticalMerge = tableCell->GetCellFormat()->GetVerticalMerge();
        short horizontalMerge = tableCell->GetGridSpan();

        if (verticalMerge == CellMerge::None && horizontalMerge == 1)
        {
            stringBuidler.append(L"Row " + std::to_wstring(i) + L", cell " + std::to_wstring(j) + L": ");
            stringBuidler.append(L"This cell isn't merged.\n");
        }
        else
        {
            stringBuidler.append(L"Row " + std::to_wstring(i) + L", cell " + std::to_wstring(j) + L": ");
            stringBuidler.append(L"This cell is merged.\n");
        }
    }
    stringBuidler.append(L"\n");
}
```

---

# spire.doc cpp table row
## clone a table row in a document
```cpp
//Get the first section
intrusive_ptr<Section> se = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first row of the first table
intrusive_ptr<TableRow> firstRow = Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->GetItemInRowCollection(0);

//Copy the first row to clone_FirstRow via TableRow.clone()
intrusive_ptr<TableRow> clone_FirstRow = firstRow->CloneTableRow();

Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0))->GetRows()->Add(clone_FirstRow);
```

---

# spire.doc cpp table
## clone and modify table in document
```cpp
//Get the first section
intrusive_ptr<Section> se = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table
intrusive_ptr<Table> original_Table = Object::Dynamic_cast<Table>(se->GetTables()->GetItemInTableCollection(0));

//Copy the existing table to copied_Table via Table.clone()
intrusive_ptr<Table> copied_Table = original_Table->CloneTable();

//Get the last row of table
intrusive_ptr<TableRow> lastRow = copied_Table->GetRows()->GetItemInRowCollection(copied_Table->GetRows()->GetCount() - 1);
//Change last row data
for (int i = 0; i < lastRow->GetCells()->GetCount() - 1; i++)
{
    lastRow->GetCells()->GetItemInCellCollection(i)->GetParagraphs()->GetItemInParagraphCollection(0)->SetText(L"New text");
}
//Add copied_Table in section
se->GetTables()->Add(copied_Table);
```

---

# spire.doc cpp table
## split table into two tables
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//We will split the table at the third row;
int splitIndex = 2;

//Create a new table for the split table
intrusive_ptr<Table> newTable = new Table(section->GetDocument());

//Add rows to the new table
for (int i = splitIndex; i < table->GetRows()->GetCount(); i++)
{
	newTable->GetRows()->Add(table->GetRows()->GetItemInRowCollection(i)->CloneTableRow());
}

//Remove rows from original table
for (int i = table->GetRows()->GetCount() - 1; i >= splitIndex; i--)
{
	table->GetRows()->RemoveAt(i);
}

//Add the new table in section
section->GetTables()->Add(newTable);
```

---

# spire.doc cpp table
## create nested table in word document
```cpp
using namespace Spire::Doc;

int main()
{
	//Create a new document
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> section = doc->AddSection();

	//Add a table
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(2, 2);

	//Set column width
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->AutoFit(AutoFitBehaviorType::AutoFitToWindow);

	//Insert content to cells
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"Spire.Doc for .NET");
	table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Description text");

	//Add a nested table to cell(first row, second column)
	intrusive_ptr<Table> nestedTable = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddTable(true);
	nestedTable->ResetCells(4, 3);
	nestedTable->AutoFit(AutoFitBehaviorType::AutoFitToContents);

	//Add content to nested cells
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"NO.");
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Item");
	nestedTable->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"Price");

	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"1");
	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Pro Edition");
	nestedTable->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$799");

	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"2");
	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Standard Edition");
	nestedTable->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$599");

	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"3");
	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"Free Edition");
	nestedTable->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(2)->AddParagraph()->AppendText(L"$0");
}
```

---

# spire.doc cpp table
## create and format table in word document
```cpp
void addTable(intrusive_ptr<Section> section)
{
	intrusive_ptr<Table> table = section->AddTable(true);
	table->ResetCells(data.size() + 1, header.size());

	// ***************** First Row *************************
	intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(0);
	row->SetIsHeader(true);
	row->SetHeight(20); //unit: point, 1point = 0.3528 mm
	row->SetHeightType(TableRowHeightType::Exactly);

	for (int i = 0; i < row->GetCells()->GetCount(); i++)
	{
		row->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetGray());
	}

	for (size_t i = 0; i < header.size(); i++)
	{
		row->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
		intrusive_ptr<Paragraph> p = row->GetCells()->GetItemInCellCollection(i)->AddParagraph();
		p->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
		intrusive_ptr<TextRange> txtRange = p->AppendText(header[i].c_str());
		txtRange->GetCharacterFormat()->SetBold(true);
	}

	for (size_t r = 0; r < data.size(); r++)
	{
		intrusive_ptr<TableRow> dataRow = table->GetRows()->GetItemInRowCollection(r + 1);
		dataRow->SetHeight(20);
		dataRow->SetHeightType(TableRowHeightType::Exactly);

		for (int i = 0; i < dataRow->GetCells()->GetCount(); i++)
		{
			dataRow->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::Empty());
		}


		for (size_t c = 0; c < data[r].size(); c++)
		{
			dataRow->GetCells()->GetItemInCellCollection(c)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
			dataRow->GetCells()->GetItemInCellCollection(c)->AddParagraph()->AppendText(data[r][c].c_str());
		}
	}

	for (int j = 1; j < table->GetRows()->GetCount(); j++)
	{
		if (j % 2 == 0)
		{
			intrusive_ptr<TableRow> row2 = table->GetRows()->GetItemInRowCollection(j);
			for (int f = 0; f < row2->GetCells()->GetCount(); f++)
			{
				row2->GetCells()->GetItemInCellCollection(f)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetLightBlue());
			}
		}
	}
}
```

---

# spire.doc cpp table
## create table directly in word document
```cpp
//Create a Word document
intrusive_ptr<Document> doc = new Document();

//Add a section
intrusive_ptr<Section> section = doc->AddSection();

//Create a table 
intrusive_ptr<Table> table = new Table(doc);
table->ResetCells(1, 2);
//Set the width of table
table->SetPreferredWidth(new PreferredWidth(WidthType::Percentage, 100));
//Set the border of table
table->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);

//Create a table row
intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(0);
row->SetHeight(50.0f);

//Create a table cell
intrusive_ptr<TableCell> cell1 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);
//Add a paragraph
intrusive_ptr<Paragraph> para1 = cell1->AddParagraph();
//Append text in the paragraph
para1->AppendText(L"Row 1, Cell 1");
//Set the horizontal alignment of paragrah
para1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
//Set the background color of cell
cell1->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());
//Set the vertical alignment of paragraph
cell1->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);

//Create a table cell
intrusive_ptr<TableCell> cell2 = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1);
intrusive_ptr<Paragraph> para2 = cell2->AddParagraph();
para2->AppendText(L"Row 1, Cell 2");
para2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
cell2->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetCadetBlue());
cell2->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
row->GetCells()->Add(cell2);

//Add the table in the section
section->GetTables()->Add(table);
```

---

# spire.doc cpp table from html
## create table in Word document from HTML string
```cpp
using namespace Spire::Doc;

//Create a Word document
intrusive_ptr<Document> document = new Document();

//Add a section
intrusive_ptr<Section> section = document->AddSection();

//Add a paragraph and append html string
section->AddParagraph()->AppendHTML(htmlContent.c_str());
```

---

# spire.doc cpp table
## create vertical table in Word document
```cpp
//Add a table with rows and columns and set the text for the table.
intrusive_ptr<Table> table = section->AddTable();
table->ResetCells(1, 1);
intrusive_ptr<TableCell> cell = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0);
table->GetRows()->GetItemInRowCollection(0)->SetHeight(150);
cell->AddParagraph()->AppendText(L"Draft copy in vertical style");

//Set the TextDirection for the table to RightToLeftRotated.
cell->GetCellFormat()->SetTextDirection(TextDirection::RightToLeftRotated);

//Set the table format.
table->GetFormat()->SetWrapTextAround(true);
table->GetFormat()->GetPositioning()->SetVertRelationTo(VerticalRelation::Page);
table->GetFormat()->GetPositioning()->SetHorizRelationTo(HorizontalRelation::Page);
table->GetFormat()->GetPositioning()->SetHorizPosition(section->GetPageSetup()->GetPageSize()->GetWidth() - table->GetWidth());
table->GetFormat()->GetPositioning()->SetVertPosition(200);
```

---

# spire.doc cpp table borders
## set different borders for table and cells
```cpp
void setTableBorders(intrusive_ptr<Table> table)
{
	table->GetFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetFormat()->GetBorders()->SetLineWidth(3.0F);
	table->GetFormat()->GetBorders()->SetColor(Color::GetRed());
}

void setCellBorders(intrusive_ptr<TableCell> tableCell)
{
	tableCell->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::DotDash);
	tableCell->GetCellFormat()->GetBorders()->SetLineWidth(1.0F);
	tableCell->GetCellFormat()->GetBorders()->SetColor(Color::GetGreen());
}

//Set borders of table
setTableBorders(table);

//Set borders of cell
setCellBorders(table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0));
```

---

# spire.doc cpp table formatting
## format merged cells in word table
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Add a new section
intrusive_ptr<Section> section = document->AddSection();

//Add a new table
intrusive_ptr<Table> table = section->AddTable(true);
table->ResetCells(4, 3);

//Create a new style
intrusive_ptr<ParagraphStyle> style = new ParagraphStyle(document);
style->SetName(L"Style");
style->GetCharacterFormat()->SetTextColor(Color::GetDeepSkyBlue());
style->GetCharacterFormat()->SetItalic(true);
style->GetCharacterFormat()->SetBold(true);
style->GetCharacterFormat()->SetFontSize(13);
document->GetStyles()->Add(style);

//Merge cell horizontally
table->ApplyHorizontalMerge(0, 0, 1);
//Apply style
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->ApplyStyle(style->GetName());
//Set vertical and horizontal alignment
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

//Merge cell vertically
table->ApplyVerticalMerge(0, 1, 3);
//Apply style
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->ApplyStyle(style->GetName());
//Set vertical and horizontal alignment
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
//Set column width
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(20, CellWidthType::Percentage);
```

---

# spire.doc c++ table diagonal border
## get diagonal border properties from table cell
```cpp
std::wstring GetBorderStyle(BorderStyle value)
{
	switch (value)
	{
	case Spire::Doc::BorderStyle::None:
		return L"None";
		break;
	case Spire::Doc::BorderStyle::Single:
		return L"Single";
		break;
	case Spire::Doc::BorderStyle::Thick:
		return L"Thick";
		break;
	case Spire::Doc::BorderStyle::Double:
		return L"Double";
		break;
	case Spire::Doc::BorderStyle::Hairline:
		return L"Hairline";
		break;
	case Spire::Doc::BorderStyle::Dot:
		return L"Dot";
		break;
	case Spire::Doc::BorderStyle::DashLargeGap:
		return L"DashLargeGap";
		break;
	case Spire::Doc::BorderStyle::DotDash:
		return L"DotDash";
		break;
	case Spire::Doc::BorderStyle::DotDotDash:
		return L"DotDotDash";
		break;
	case Spire::Doc::BorderStyle::Triple:
		return L"Triple";
		break;
	case Spire::Doc::BorderStyle::ThinThickSmallGap:
		return L"ThinThickSmallGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinSmallGap:
		return L"ThickThinSmallGap";
		break;

	case Spire::Doc::BorderStyle::ThinThickThinSmallGap:
		return L"ThinThickThinSmallGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickMediumGap:
		return L"ThinThickMediumGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinMediumGap:
		return L"ThickThinMediumGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickThinMediumGap:
		return L"ThinThickThinMediumGap";
		break;

	case Spire::Doc::BorderStyle::ThinThickLargeGap:
		return L"ThinThickLargeGap";
		break;
	case Spire::Doc::BorderStyle::ThickThinLargeGap:
		return L"ThickThinLargeGap";
		break;
	case Spire::Doc::BorderStyle::ThinThickThinLargeGap:
		return L"ThinThickThinLargeGap";
		break;
	case Spire::Doc::BorderStyle::Wave:
		return L"Wave";
		break;
	case Spire::Doc::BorderStyle::DoubleWave:
		return L"DoubleWave";
		break;
	case Spire::Doc::BorderStyle::DashSmallGap:
		return L"DashSmallGap";
		break;
	case Spire::Doc::BorderStyle::DashDotStroker:
		return L"DashDotStroker";
		break;
	case Spire::Doc::BorderStyle::Emboss3D:
		return L"Emboss3D";
		break;
	case Spire::Doc::BorderStyle::Engrave3D:
		return L"Engrave3D";
		break;
	case Spire::Doc::BorderStyle::Outset:
		return L"Outset";
		break;
	case Spire::Doc::BorderStyle::Inset:
		return L"Inset";
		break;
	case Spire::Doc::BorderStyle::TwistedLines1:
		return L"TwistedLines1";
		break;
	case Spire::Doc::BorderStyle::Cleared:
		return L"Cleared";
		break;
	}
	return L"";
}

//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table in the section
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Get the setting of the diagonal border of table cell
BorderStyle bs_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetBorderType();

intrusive_ptr<Color> color_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetColor();

float width_UP = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalUp()->GetLineWidth();

BorderStyle bs_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetBorderType();

intrusive_ptr<Color> color_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetColor();

float width_Down = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetDiagonalDown()->GetLineWidth();
```

---

# spire.doc cpp table
## get table, row and cell indices
```cpp
//Get the first section
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Get the first table in the section
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Get table collections
intrusive_ptr<TableCollection> collections = section->GetTables();

//Get the table index
int tableIndex = collections->IndexOf(table);

//Get the index of the last table row
intrusive_ptr<TableRow> row = table->GetLastRow();
int rowIndex = row->GetRowIndex();

//Get the index of the last table cell
intrusive_ptr<TableCell> cell = Object::Dynamic_cast<TableCell>(row->GetLastChild());
int cellIndex = cell->GetCellIndex();
```

---

# spire.doc cpp table position
## get table position from word document
```cpp
std::wstring GetHorizontalPositionType(HorizontalPosition value)
{
	switch (value)
	{
	case Spire::Doc::HorizontalPosition::None:
		return L"None";
		break;
	case Spire::Doc::HorizontalPosition::Left:
		return L"Left";
		break;
	case Spire::Doc::HorizontalPosition::Center:
		return L"Center";
		break;
	case Spire::Doc::HorizontalPosition::Right:
		return L"Right";
		break;
	case Spire::Doc::HorizontalPosition::Inside:
		return L"Inside";
		break;
	case Spire::Doc::HorizontalPosition::Outside:
		return L"Outside";
		break;
	case Spire::Doc::HorizontalPosition::Inline:
		return L"Inline";
		break;
	}
	return L"";
}
std::wstring GetHorizontalRelationType(HorizontalRelation value)
{
	switch (value)
	{
	case Spire::Doc::HorizontalRelation::Column:
		return L"Column";
		break;
	case Spire::Doc::HorizontalRelation::Margin:
		return L"Margin";
		break;
	case Spire::Doc::HorizontalRelation::Page:
		return L"Page";
		break;
	}
	return L"";
}
std::wstring GetVerticalPositionType(VerticalPosition value)
{
	switch (value)
	{
	case Spire::Doc::VerticalPosition::None:
		return L"None";
		break;
	case Spire::Doc::VerticalPosition::Top:
		return L"Top";
		break;
	case Spire::Doc::VerticalPosition::Center:
		return L"Center";
		break;
	case Spire::Doc::VerticalPosition::Bottom:
		return L"Bottom";
		break;
	case Spire::Doc::VerticalPosition::Inside:
		return L"Inside";
		break;
	case Spire::Doc::VerticalPosition::Outside:
		return L"Outside";
		break;
	case Spire::Doc::VerticalPosition::Inline:
		return L"Inline";
		break;
	}
	return L"";
}
std::wstring GetVerticalRelationType(VerticalRelation value)
{
	switch (value)
	{
	case Spire::Doc::VerticalRelation::Margin:
		return L"Margin";
		break;
	case Spire::Doc::VerticalRelation::Page:
		return L"Page";
		break;
	case Spire::Doc::VerticalRelation::Paragraph:
		return L"Paragraph";
		break;
	}
	return L"";
}

//Create a document
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
//Get the first table
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Verify whether the table uses "Around" text wrapping or not.
if (table->GetFormat()->GetWrapTextAround())
{
	intrusive_ptr<TablePositioning> positon = table->GetFormat()->GetPositioning();

	// Get horizontal position information
	positon->GetHorizPosition();
	GetHorizontalPositionType(positon->GetHorizPositionAbs());
	GetHorizontalRelationType(positon->GetHorizRelationTo());
	
	// Get vertical position information
	positon->GetVertPosition();
	GetVerticalPositionType(positon->GetVertPositionAbs());
	GetVerticalRelationType(positon->GetVertRelationTo());
	
	// Get distance from surrounding text
	positon->GetDistanceFromTop();
	positon->GetDistanceFromLeft();
	positon->GetDistanceFromBottom();
	positon->GetDistanceFromRight();
}

document->Close();
```

---

# spire.doc cpp helloworld
## Create a simple Hello World document
```cpp
using namespace Spire::Doc;

//Create word document
intrusive_ptr<Document> document = new Document();

//Create a new section
intrusive_ptr<Section> section = document->AddSection();

//Create a new paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Append Text
paragraph->AppendText(L"Hello World!");
```

---

# spire.doc cpp table
## merge and split table cells
```cpp
//Merge cells horizontally
table->ApplyHorizontalMerge(6, 2, 3);

//Merge cells vertically
table->ApplyVerticalMerge(2, 4, 5);

//Split the cell
table->GetRows()->GetItemInRowCollection(8)->GetCells()->GetItemInCellCollection(3)->SplitCell(2, 2);
```

---

# Spire.Doc Table Format Modification
## Modify table, row, and cell formatting in Word documents
```cpp
void MoidyTableFormat(intrusive_ptr<Table> table)
{
    //Set table width
    table->SetPreferredWidth(new PreferredWidth(WidthType::Twip, static_cast<short>(6000)));

    //Apply style for table
    table->ApplyStyle(DefaultTableStyle::ColorfulGridAccent3);

    //Set table padding
    table->GetFormat()->GetPaddings()->SetAll(5);

    //Set table title and description
    table->SetTitle(L"Spire.Doc for .NET");
    table->SetTableDescription(L"Spire.Doc for .NET is a professional Word .NET library");
}

void ModifyRowFormat(intrusive_ptr<Table> table)
{
    //Set cell spacing
    table->GetFormat()->SetCellSpacing(2);

    //Set row height
    table->GetRows()->GetItemInRowCollection(1)->SetHeightType(TableRowHeightType::Exactly);
    table->GetRows()->GetItemInRowCollection(1)->SetHeight(20.0f);

    //Set background color
    for (int i = 0; i < table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetCount(); i++)
    {
        table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(i)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetDarkSeaGreen());
    }
}

void ModifyCellFormat(intrusive_ptr<Table> table)
{
    //Set alignment
    table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
    table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

    //Set background color
    table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetShading()->SetBackgroundPatternColor(Color::GetDarkSeaGreen());

    //Set cell border
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->SetLineWidth(1.0f);
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetLeft()->SetColor(Color::GetRed());
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetTop()->SetColor(Color::GetRed());
    table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->GetBorders()->GetBottom()->SetColor(Color::GetRed());

    //Set text direction
    table->GetRows()->GetItemInRowCollection(3)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetTextDirection(TextDirection::RightToLeft);
}
```

---

# spire.doc cpp table
## prevent page breaks in table
```cpp
//Get the table from Word document.
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(document->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

//Change the paragraph setting to keep them together.
for (int i = 0; i < table->GetRows()->GetCount(); i++)
{
    intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
    for (int j = 0; j < row->GetCells()->GetCount(); j++)
    {
        intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(j);
        for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
        {
            intrusive_ptr<Paragraph> p = cell->GetParagraphs()->GetItemInParagraphCollection(k);
            p->GetFormat()->SetKeepFollow(true);
        }
    }
}
```

---

# spire.doc cpp table
## remove table from document
```cpp
//Remove the first Table            
doc->GetSections()->GetItemInSectionCollection(0)->GetTables()->RemoveAt(0);
```

---

# spire.doc cpp table
## repeat header rows on each page
```cpp
//Create a table with default borders
intrusive_ptr<Table> table = section->AddTable(true);

//Add a header row that will repeat on each page
intrusive_ptr<TableRow> row = table->AddRow();
//Set the row as a table header (this makes it repeat on each page)
row->SetIsHeader(true);

//Add a cell to the header row
intrusive_ptr<TableCell> cell = row->AddCell();
//Add paragraph and text to the cell
intrusive_ptr<Paragraph> paragraph = cell->AddParagraph();
paragraph->AppendText(L"Row Header 1");

//Add a second header row that will also repeat on each page
row = table->AddRow(false, 1);
row->SetIsHeader(true); // This makes it repeat on each page too

//Add a cell to the second header row
cell = row->GetCells()->GetItemInCellCollection(0);
//Add paragraph and text to the cell
paragraph = cell->AddParagraph();
paragraph->AppendText(L"Row Header 2");

//Add regular rows (these won't repeat on each page)
for (int i = 0; i < 70; i++)
{
    row = table->AddRow(false, 2);
    cell = row->GetCells()->GetItemInCellCollection(0);
    cell->AddParagraph()->AppendText(L"Column 1 Text");
    cell = row->GetCells()->GetItemInCellCollection(1);
    cell->AddParagraph()->AppendText(L"Column 2 Text");
}
```

---

# spire.doc cpp table text replacement
## replace text in table using regex and string matching
```cpp
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
```

---

# spire.doc cpp table
## set column width for table
```cpp
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Traverse the first column
for (int i = 0; i < table->GetRows()->GetCount(); i++)
{
    //Set the width and type of the cell
    table->GetRows()->GetItemInRowCollection(i)->GetCells()->GetItemInCellCollection(0)->SetCellWidth(200, CellWidthType::Point);
}
```

---

# spire.doc cpp table positioning
## set table outside position relative to image in document header
```cpp
//Create a new word document and add new section
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> sec = doc->AddSection();

//Get header
intrusive_ptr<HeaderFooter> header = doc->GetSections()->GetItemInSectionCollection(0)->GetHeadersFooters()->GetHeader();

//Add new paragraph on header and set HorizontalAlignment of the paragraph as left
intrusive_ptr<Paragraph> paragraph = header->AddParagraph();
paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

//Add a table of 4 rows and 2 columns
intrusive_ptr<Table> table = header->AddTable();
table->ResetCells(4, 2);

//Set the position of the table to the right of the image
table->GetFormat()->SetWrapTextAround(true);
table->GetFormat()->GetPositioning()->SetHorizPositionAbs(HorizontalPosition::Outside);
table->GetFormat()->GetPositioning()->SetVertRelationTo(VerticalRelation::Margin);
table->GetFormat()->GetPositioning()->SetVertPosition(43);
```

---

# spire.doc cpp table
## set table style and border
```cpp
//Get the first table
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

//Apply the table style
table->ApplyStyle(DefaultTableStyle::ColorfulList);

//Set right border of table
table->GetFormat()->GetBorders()->GetRight()->SetBorderType(BorderStyle::Hairline);
table->GetFormat()->GetBorders()->GetRight()->SetLineWidth(1.0F);
table->GetFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());

//Set top border of table
table->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Hairline);
table->GetFormat()->GetBorders()->GetTop()->SetLineWidth(1.0F);
table->GetFormat()->GetBorders()->GetTop()->SetColor(Color::GetGreen());

//Set left border of table
table->GetFormat()->GetBorders()->GetLeft()->SetBorderType(BorderStyle::Hairline);
table->GetFormat()->GetBorders()->GetLeft()->SetLineWidth(1.0F);
table->GetFormat()->GetBorders()->GetLeft()->SetColor(Color::GetYellow());

//Set bottom border is none
table->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::DotDash);

//Set vertical and horizontal border 
table->GetFormat()->GetBorders()->GetVertical()->SetBorderType(BorderStyle::Dot);
table->GetFormat()->GetBorders()->GetHorizontal()->SetBorderType(BorderStyle::None);
table->GetFormat()->GetBorders()->GetVertical()->SetColor(Color::GetOrange());
```

---

# spire.doc cpp table vertical alignment
## set vertical alignment for table cells in Word document
```cpp
//Create a new Word document and add a new section
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> section = doc->AddSection();

//Add a table with 3 columns and 3 rows
intrusive_ptr<Table> table = section->AddTable(true);
table->ResetCells(3, 3);

//Merge rows
table->ApplyVerticalMerge(0, 0, 2);

//Set the vertical alignment for each cell, default is top
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);
table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);
```

---

# spire.doc cpp image hyperlink
## create image hyperlink in document
```cpp
//Load Document
intrusive_ptr<Document> doc = new Document();

intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
//Add a paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
//Load an image to a DocPicture object
#if defined(SKIASHARP)
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
//Add an image hyperlink to the paragraph
picture->LoadImageSpire(inputFile.c_str()_1);
#else
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
//Add an image hyperlink to the paragraph
picture->LoadImageSpire(inputFile_1.c_str());
#endif
paragraph->AppendHyperlink(L"https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType::WebLink);
```

---

# Spire.Doc C++ Hyperlink Extraction
## Find and extract hyperlinks from a Word document
```cpp
//Create a hyperlink list
std::vector<intrusive_ptr<Field>> hyperlinks;
std::wstring hyperlinksText = L"";
//Iterate through the items in the sections to find all hyperlinks
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> docObj = section->GetBody()->GetChildObjects()->GetItem(j);
        if (docObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
        {
            intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(docObj);
            for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
            {
                intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
                if (obj->GetDocumentObjectType() == DocumentObjectType::Field)
                {
                    intrusive_ptr<Field> field = Object::Dynamic_cast<Field>(obj);
                    if (field->GetType() == FieldType::FieldHyperlink)
                    {
                        hyperlinks.push_back(field);
                        //Get the hyperlink text
                        std::wstring text = field->GetFieldText();
                        hyperlinksText.append(text.append(L"\n"));
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc C++ Hyperlink
## Insert different types of hyperlinks in a Word document
```cpp
void InsertHyperlink(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();
	paragraph->AppendText(L"Spire.Doc for .NET \n e-iceblue company Ltd. 2002-2010 All rights reserverd");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Home page");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Contact US");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"mailto:support@e-iceblue.com", L"support@e-iceblue.com", HyperlinkType::EMailLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Forum");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/forum/", L"www.e-iceblue.com/forum/", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Download Link");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", L"www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Insert Link On Image");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(DataPath"/Demo/Spire.Doc.png");
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType::WebLink);
}
```

---

# spire.doc cpp hyperlink
## modify hyperlink text in document
```cpp
intrusive_ptr<Document> doc = new Document();

//Find all hyperlinks in the document
std::vector<intrusive_ptr<Field>> hyperlinks;
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
	for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
	{
		intrusive_ptr<DocumentObject> docObj = section->GetBody()->GetChildObjects()->GetItem(j);
		if (docObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
		{
			intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(docObj);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
				if (obj->GetDocumentObjectType() == DocumentObjectType::Field)
				{
					intrusive_ptr<Field> field = Object::Dynamic_cast<Field>(obj);
					if (field->GetType() == FieldType::FieldHyperlink)
					{
						hyperlinks.push_back(field);
					}
				}
			}
		}
	}
}

//Reset the property of hyperlinks->GetItem(0)->FieldText by using the index of the hyperlink
hyperlinks[0]->SetFieldText(L"Spire.Doc component");
```

---

# spire.doc c++ hyperlinks
## remove hyperlinks from document
```cpp
std::vector<intrusive_ptr<Field>> FindAllHyperlinks(intrusive_ptr<Document> document)
{
	std::vector<intrusive_ptr<Field>> hyperlinks;
	//Iterate through the items in the sections to find all hyperlinks
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		int secBodyChildCount = section->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secBodyChildCount; j++)
		{
			intrusive_ptr<DocumentObject> childObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (childObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				int paraChildCount = (Object::Dynamic_cast<Paragraph>(childObj))->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildCount; k++)
				{
					intrusive_ptr<DocumentObject> paraObj = (Object::Dynamic_cast<Paragraph>(childObj))->GetChildObjects()->GetItem(k);
					if (paraObj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						intrusive_ptr<Field> field = Object::Dynamic_cast<Field>(paraObj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
						}
					}
				}
			}
		}
	}
	return hyperlinks;
}

void FlattenHyperlinks(intrusive_ptr<Field> field)
{
	int ownerParaIndex = field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetOwnerParagraph());
	int fieldIndex = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
	intrusive_ptr<Paragraph> sepOwnerPara = field->GetSeparator()->GetOwnerParagraph();
	int sepOwnerParaIndex = field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetSeparator()->GetOwnerParagraph());
	int sepIndex = field->GetSeparator()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetSeparator());
	int endIndex = field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetEnd());
	int endOwnerParaIndex = field->GetEnd()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetEnd()->GetOwnerParagraph());

	FormatFieldResultText(field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody(), sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

	field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->RemoveAt(endIndex);

	for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
	{
		if (i == sepOwnerParaIndex && i == ownerParaIndex)
		{
			for (int j = sepIndex; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);

			}
		}
		else if (i == ownerParaIndex)
		{
			for (int j = field->GetOwnerParagraph()->GetChildObjects()->GetCount() - 1; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex; j >= 0; j--)
			{
				sepOwnerPara->GetChildObjects()->RemoveAt(j);
			}
		}
		else
		{
			field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->RemoveAt(i);
		}
	}
}

void FormatFieldResultText(intrusive_ptr<Body> ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
{
	for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
	{
		intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(ownerBody->GetChildObjects()->GetItem(i));
		if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < endIndex; j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else if (i == endOwnerParaIndex)
		{
			for (int j = 0; j < endIndex; j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else
		{
			for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
	}
}

void FormatText(intrusive_ptr<TextRange> tr)
{
	//Set the text color to black
	tr->GetCharacterFormat()->SetTextColor(Color::GetBlack());
	//Set the text underline style to none
	tr->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::None);
}
```

---

# spire.doc cpp hyperlink formatting
## create and format hyperlinks in document
```cpp
//Load Document
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());
intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

//Add a paragraph and append a hyperlink to the paragraph
intrusive_ptr<Paragraph> para1 = section->AddParagraph();
para1->AppendText(L"Regular Link: ");
//Format the hyperlink with default color and underline style
intrusive_ptr<TextRange> txtRange1 = para1->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
txtRange1->GetCharacterFormat()->SetFontName(L"Times New Roman");
txtRange1->GetCharacterFormat()->SetFontSize(12);
intrusive_ptr<Paragraph> blankPara1 = section->AddParagraph();

//Add a paragraph and append a hyperlink to the paragraph
intrusive_ptr<Paragraph> para2 = section->AddParagraph();
para2->AppendText(L"Change Color: ");
//Format the hyperlink with red color and underline style
intrusive_ptr<TextRange> txtRange2 = para2->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
txtRange2->GetCharacterFormat()->SetFontName(L"Times New Roman");
txtRange2->GetCharacterFormat()->SetFontSize(12);
txtRange2->GetCharacterFormat()->SetTextColor(Color::GetRed());
intrusive_ptr<Paragraph> blankPara2 = section->AddParagraph();

//Add a paragraph and append a hyperlink to the paragraph
intrusive_ptr<Paragraph> para3 = section->AddParagraph();
para3->AppendText(L"Remove Underline: ");
//Format the hyperlink with red color and no underline style
intrusive_ptr<TextRange> txtRange3 = para3->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
txtRange3->GetCharacterFormat()->SetFontName(L"Times New Roman");
txtRange3->GetCharacterFormat()->SetFontSize(12);
txtRange3->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::None);
```

---

# spire.doc cpp decrypt
## decrypt password protected word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str(), FileFormat::Docx, L"E-iceblue");
```

---

# spire.doc cpp security
## encrypt document with password
```cpp
//Encrypt document with password specified by textBox1
document->Encrypt(L"E-iceblue");
```

---

# spire.doc cpp security
## lock specified sections in document
```cpp
//Protect the document with AllowOnlyFormFields protection type.
document->Protect(ProtectionType::AllowOnlyFormFields, L"123");

//Unprotect section 2
s2->SetProtectForm(false);
```

---

# spire.doc cpp security
## remove editable ranges from document
```cpp
//Find "PermissionStart" and "PermissionEnd" tags and remove them
for (int i = 0; i < document->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetBody()->GetParagraphs()->GetCount(); j++)
    {
        intrusive_ptr<Paragraph> para = section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(j);
        for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
        {
            intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
            if (Object::CheckType<PermissionStart>(obj) || Object::CheckType<PermissionEnd>(obj))
            {
                para->GetChildObjects()->Remove(obj);
            }
            else
            {
                k++;
            }
        }
    }
}
```

---

# Spire.Doc C++ Security
## Remove read-only restriction from document
```cpp
intrusive_ptr<Document> doc = new Document();
doc->LoadFromFile(inputFile.c_str());
//Remove ReadOnly Restriction.
doc->Protect(ProtectionType::NoProtection);
doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
doc->Close();
```

---

# spire.doc cpp security
## set editable range in protected document
```cpp
//Create a new document
intrusive_ptr<Document> document = new Document();
//Protect whole document
document->Protect(ProtectionType::AllowOnlyReading, L"password");
//Create tags for permission start and end
intrusive_ptr<PermissionStart> start = new PermissionStart(document, L"testID");
intrusive_ptr<PermissionEnd> end = new PermissionEnd(document, L"testID");
//Add the start and end tags to allow the first paragraph to be edited.
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Insert(0, start);
document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Add(end);
document->Close();
```

---

# Spire.Doc Document Protection
## Apply specified protection type to Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Protect the Word file.
document->Protect(ProtectionType::AllowOnlyReading, L"123456");
```

---

# spire.doc cpp security
## convert Word document to encrypted PDF
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Create an instance of ToPdfParameterList.
intrusive_ptr<ToPdfParameterList> toPdf = new ToPdfParameterList();

//Set the user password for the resulted PDF file.
toPdf->GetPdfSecurity()->Encrypt(L"e-iceblue");

//Save to file.
document->SaveToFile(outputFile.c_str(), toPdf);
document->Close();
```

---

# Spire.Doc C++ Fields
## Add TC (Table of Contents Entry) field to Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Add a new section.
intrusive_ptr<Section> section = document->AddSection();

//Add a new paragraph.
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Add TC field in the paragraph
intrusive_ptr<Field> field = paragraph->AppendField(L"TC", FieldType::FieldTOCEntry);
wstring codeString = L"TC ";
codeString += L"\"Entry Text\"";
codeString += L" \\f";
codeString += L" t";
field->SetCode(codeString.c_str());
```

---

# spire.doc cpp fields conversion
## convert form fields to body text in document
```cpp
//Create the source document
intrusive_ptr<Document> sourceDocument = new Document();

//Traverse FormFields
int formFieldsCount = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount();
for (int i = 0; i < formFieldsCount; i++)
{
    intrusive_ptr<FormField> field = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);
    //Find FieldFormTextInput type field
    if (field->GetType() == FieldType::FieldFormTextInput)
    {
        //Get the paragraph
        intrusive_ptr<Paragraph> paragraph = field->GetOwnerParagraph();

        //Define variables
        int startIndex = 0;
        int endIndex = 0;

        //Create a new TextRange
        intrusive_ptr<TextRange> textRange = new TextRange(sourceDocument);

        //Set text for textRange
        textRange->SetText(paragraph->GetText());

        //Traverse DocumentObjectS of field paragraph
        int pChildObjectsCount = paragraph->GetChildObjects()->GetCount();
        for (int j = 0; j < pChildObjectsCount; j++)
        {
            intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(j);
            //If its DocumentObjectType is BookmarkStart
            if (obj->GetDocumentObjectType() == DocumentObjectType::BookmarkStart)
            {
                //Get the index
                startIndex = paragraph->GetChildObjects()->IndexOf(obj);
            }
            //If its DocumentObjectType is BookmarkEnd
            if (obj->GetDocumentObjectType() == DocumentObjectType::BookmarkEnd)
            {
                //Get the index
                endIndex = paragraph->GetChildObjects()->IndexOf(obj);
            }
        }
        //Remove ChildObjects
        for (int k = endIndex; k > startIndex; k--)
        {
            //If it is TextFormField
            if (Object::CheckType<TextFormField>(paragraph->GetChildObjects()->GetItem(k)))
            {
                intrusive_ptr<TextFormField> textFormField = boost::dynamic_pointer_cast<TextFormField>(paragraph->GetChildObjects()->GetItem(k));

                //Remove the field object
                paragraph->GetChildObjects()->Remove(textFormField);
            }
            else
            {
                paragraph->GetChildObjects()->RemoveAt(k);
            }
        }
        //Insert the new TextRange
        paragraph->GetChildObjects()->Insert(startIndex, textRange);

        break;
    }
}
```

---

# spire.doc cpp fields
## convert document fields to text
```cpp
//Get all fields in document
intrusive_ptr<FieldCollection> fields = document->GetFields();
int count = fields->GetCount();

for (int i = 0; i < count; i++)
{
    intrusive_ptr<Field> field = fields->GetItem(0);
    std::wstring s = field->GetFieldText();
    int index = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
    intrusive_ptr<TextRange> textRange = new TextRange(document);
    textRange->SetText(s.c_str());
    textRange->GetCharacterFormat()->SetFontSize(24.0f);

    field->GetOwnerParagraph()->GetChildObjects()->Insert(index, textRange);
    field->GetOwnerParagraph()->GetChildObjects()->Remove(field);
}
```

---

# spire.doc cpp field conversion
## convert IF fields to text in document
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Get all fields in document
intrusive_ptr<FieldCollection> fields = document->GetFields();

for (int i = 0; i < fields->GetCount(); i++)
{
	intrusive_ptr<Field> field = fields->GetItem(i);
	if (field->GetType() == FieldType::FieldIf)
	{
		intrusive_ptr<TextRange> original = Object::Dynamic_cast<TextRange>(field);
		//Get field text
		std::wstring text = field->GetFieldText();
		//Create a new textRange and set its format
		intrusive_ptr<TextRange> textRange = new TextRange(document);
		textRange->SetText(text.c_str());
		textRange->GetCharacterFormat()->SetFontName(original->GetCharacterFormat()->GetFontName());
		textRange->GetCharacterFormat()->SetFontSize(original->GetCharacterFormat()->GetFontSize());

		intrusive_ptr<Paragraph> par = field->GetOwnerParagraph();
		//Get the index of the if field
		int index = par->GetChildObjects()->IndexOf(field);
		//Remove if field via index
		par->GetChildObjects()->RemoveAt(index);
		//Insert field text at the position of if field
		par->GetChildObjects()->Insert(index, textRange);
	}
}
```

---

# spire.doc cpp cross-reference
## create cross-reference field with bookmark in Word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Add a new section.
intrusive_ptr<Section> section = document->AddSection();

//Create a bookmark.
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
paragraph->AppendBookmarkStart(L"MyBookmark");
paragraph->AppendText(L"Text inside a bookmark");
paragraph->AppendBookmarkEnd(L"MyBookmark");

//Insert line breaks.
for (int i = 0; i < 4; i++)
{
    paragraph->AppendBreak(BreakType::LineBreak);
}

//Create a cross-reference field, and link it to bookmark.                    
intrusive_ptr<Field> field = new Field(document);
field->SetType(FieldType::FieldRef);
field->SetCode(L"REF MyBookmark \\p \\h");

//Insert field to paragraph.
paragraph = section->AddParagraph();
paragraph->AppendText(L"For more information, see ");
paragraph->GetChildObjects()->Add(field);

//Insert FieldSeparator object.
intrusive_ptr<FieldMark> fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
paragraph->GetChildObjects()->Add(fieldSeparator);

//Set display text of the field.
intrusive_ptr<TextRange> tr = new TextRange(document);
tr->SetText(L"above");
paragraph->GetChildObjects()->Add(tr);

//Insert FieldEnd object to mark the end of the field.
intrusive_ptr<FieldMark> fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
paragraph->GetChildObjects()->Add(fieldEnd);
```

---

# spire.doc cpp form fields
## create form fields in Word document
```cpp
void AddFormDemo(intrusive_ptr<Section> section)
{
    // Add table for form fields
    intrusive_ptr<Table> table = section->AddTable();
    table->SetDefaultColumnsNumber(2);
    table->SetDefaultRowHeight(20);

    // Load form config data
    tinyxml2::XMLDocument* xpathDoc = new tinyxml2::XMLDocument();
    std::wstring wpath = DATAPATH"/Form.xml";
    std::string finalPath = wstring2string(wpath);
    xpathDoc->LoadFile(finalPath.c_str());
    tinyxml2::XMLElement* root = xpathDoc->RootElement();
    
    // Process sections and fields
    for (tinyxml2::XMLElement* node = root->FirstChildElement("section"); node; node = node->NextSiblingElement("section"))
    {
        // Create a row for field group label
        intrusive_ptr<TableRow> row = table->AddRow(false);
        intrusive_ptr<Paragraph> cellParagraph = row->GetCells()->GetItemInCellCollection(0)->AddParagraph();
        cellParagraph->AppendText(string2wstring(node->Attribute("name")).c_str());

        // Process fields
        for (tinyxml2::XMLElement* fieldNode = node->FirstChildElement("field"); fieldNode; fieldNode = fieldNode->NextSiblingElement("field"))
        {
            // Create a row for field
            intrusive_ptr<TableRow> fieldRow = table->AddRow(false);

            // Field label
            intrusive_ptr<Paragraph> labelParagraph = fieldRow->GetCells()->GetItemInCellCollection(0)->AddParagraph();
            labelParagraph->AppendText(string2wstring(fieldNode->Attribute("label")).c_str());

            intrusive_ptr<Paragraph> fieldParagraph = fieldRow->GetCells()->GetItemInCellCollection(1)->AddParagraph();
            std::wstring fieldId = string2wstring(fieldNode->Attribute("id"));
            
            // Create form fields based on type
            std::wstring fieldType = string2wstring(fieldNode->Attribute("type"));
            if (fieldType == L"text")
            {
                // Add text input field
                intrusive_ptr<TextFormField> field = Object::Dynamic_cast<TextFormField>(fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormTextInput));
                field->SetDefaultText(L"");
                field->SetText(L"");
            }
            else if (fieldType == L"list")
            {
                // Add dropdown field
                intrusive_ptr<DropDownFormField> list = Object::Dynamic_cast<DropDownFormField>(fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormDropDown));

                // Add items into dropdown
                for (tinyxml2::XMLElement* itemNode = fieldNode->FirstChildElement("item"); itemNode; itemNode = itemNode->NextSiblingElement("item"))
                {
                    list->GetDropDownItems()->Add(string2wstring(itemNode->GetText()).c_str());
                }
            }
            else if (fieldType == L"checkbox")
            {
                // Add checkbox field
                fieldParagraph->AppendField(fieldId.c_str(), FieldType::FieldFormCheckBox);
            }
        }

        // Merge field group row
        table->ApplyHorizontalMerge(row->GetRowIndex(), 0, 1);
    }

    delete xpathDoc;
}
```

---

# spire.doc cpp field
## create IF field in word document
```cpp
void CreateIfField(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");

	paragraph->GetItems()->Add(ifField);
	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");

	intrusive_ptr<ParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	intrusive_ptr<FieldMark> fm = Object::Dynamic_cast<FieldMark>(end);
	fm->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}
```

---

# Spire.Doc C++ Fields
## Create Nested IF Fields in Word Document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Create an IF field
intrusive_ptr<IfField> ifField = new IfField(document);
ifField->SetType(FieldType::FieldIf);
ifField->SetCode(L"IF ");
paragraph->GetItems()->Add(ifField);

//Create the embedded IF field
intrusive_ptr<IfField> ifField2 = new IfField(document);
ifField2->SetType(FieldType::FieldIf);
ifField2->SetCode(L"IF ");
paragraph->GetChildObjects()->Add(ifField2);
paragraph->GetItems()->Add(ifField2);
paragraph->AppendText(L"\"200\" < \"50\"   \"200\" \"50\" ");
intrusive_ptr<IParagraphBase> embeddedEnd = document->CreateParagraphItem(ParagraphItemType::FieldMark);
(Object::Dynamic_cast<FieldMark>(embeddedEnd))->SetType(FieldMarkType::FieldEnd);
paragraph->GetItems()->Add(embeddedEnd);
ifField2->SetEnd(Object::Dynamic_cast<FieldMark>(embeddedEnd));

paragraph->AppendText(L" > ");
paragraph->AppendText(L"\"100\" ");
paragraph->AppendText(L"\"Thanks\" ");
paragraph->AppendText(L"\"The minimum order is 100 units\"");
intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
paragraph->GetItems()->Add(end);
ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));

//Update all fields in the document.
document->SetIsUpdateFields(true);

document->Close();
```

---

# spire.doc cpp form fields
## fill form fields in word document from xml data
```cpp
// Helper function to convert string to wstring
wstring  string2wstring(string str)
{
	std::string strLocale = setlocale(LC_ALL, "");
	const char* chSrc = str.c_str();
	size_t nDestSize = mbstowcs(NULL, chSrc, 0) + 1;
	wchar_t* wchDest = new wchar_t[nDestSize];
	wmemset(wchDest, 0, nDestSize);
	mbstowcs(wchDest, chSrc, nDestSize);
	std::wstring wstrResult = wchDest;
	delete[] wchDest;
	setlocale(LC_ALL, strLocale.c_str());
	return wstrResult;
}

// Helper function to convert wstring to string
string wstring2string(const std::wstring& wstr)
{
	std::string result;
	result.reserve(wstr.size());
	for (size_t i = 0; i < wstr.size(); ++i)
	{
		result += static_cast<char>(wstr[i] & 0xFF);
	}
	return result;
}

// Helper function to trim whitespace
wstring Trim(const std::wstring& str)
{
	auto first = find_if_not(str.begin(), str.end(), [](wint_t c) {return iswspace(c); });
	auto last = find_if_not(str.rbegin(), str.rend(), [](wint_t c) {return iswspace(c); }).base();
	return (first >= last) ? L"" : wstring(first, last);
}

// document is a loaded Word document
// user is the root element of loaded XML data
// Iterate through form fields and fill data
int formFieldsCount = document->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount();
for (int i = 0; i < formFieldsCount; i++)
{
	intrusive_ptr<FormField> field = document->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);

	tinyxml2::XMLElement* propertyNode = user->FirstChildElement(wstring2string(Trim(field->GetName())).c_str());
	if (propertyNode != nullptr)
	{
		switch (field->GetType())
		{
		case FieldType::FieldFormTextInput:
			// Fill text input field
			field->SetText(string2wstring(propertyNode->GetText()).c_str());
			break;

		case FieldType::FieldFormDropDown:
		{
			// Fill dropdown field
			intrusive_ptr<DropDownFormField> combox = Object::Dynamic_cast<DropDownFormField>(field);
			for (int j = 0; j < combox->GetDropDownItems()->GetCount(); j++)
			{
				if (combox->GetDropDownItems()->GetItem(j)->GetText() == string2wstring(propertyNode->GetText()))
				{
					combox->SetDropDownSelectedIndex(j);
					break;
				}
				if (wcscmp(field->GetName(), L"country") == 0 && wcscmp(combox->GetDropDownItems()->GetItem(j)->GetText(), L"Others") == 0)
				{
					combox->SetDropDownSelectedIndex(j);
				}
			}
			break;
		}
		case FieldType::FieldFormCheckBox:
			// Fill checkbox field
			std::string boolStr = propertyNode->GetText();
			bool value;
			std::istringstream(boolStr) >> boolalpha >> value;
			if (value)
			{
				intrusive_ptr<CheckBoxFormField> checkBox = Object::Dynamic_cast<CheckBoxFormField>(field);
				checkBox->SetChecked(true);
			}
			break;
		default:
			break;
		}
	}
}
```

---

# spire.doc cpp form fields
## modify form field properties
```cpp
//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

//Get FormField by index
intrusive_ptr<FormField> formField = section->GetBody()->GetFormFields()->GetItem(1);

if (formField->GetType() == FieldType::FieldFormTextInput)
{
	wstring formFieldName = formField->GetName();
	wstring temp = L"My name is " + formFieldName;
	formField->SetText(temp.c_str());
	formField->GetCharacterFormat()->SetTextColor(Color::GetRed());
	formField->GetCharacterFormat()->SetItalic(true);
}
```

---

# spire.doc cpp fields
## get field text from document
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();

//Get all fields in document
intrusive_ptr<FieldCollection> fields = document->GetFields();
for (int i = 0; i < fields->GetCount(); i++)
{
    intrusive_ptr<Field> field = fields->GetItem(i);
    //Get field text
    std::wstring fieldText = field->GetFieldText();
}
```

---

# spire.doc cpp form field
## get form field by name
```cpp
wstring getFormFieldType(FormFieldType type)
{
	switch (type)
	{
	case FormFieldType::CheckBox:
		return L"CheckBox";
		break;
	case FormFieldType::DropDown:
		return L"DropDown";
		break;
	case FormFieldType::TextInput:
		return L"TextInput";
		break;
	case FormFieldType::Unknown:
		return L"Unknown";
		break;
	}
	return L"";
}

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

//Get form field by name
intrusive_ptr<FormField> formField = section->GetBody()->GetFormFields()->GetItem(L"email");
wstring formFieldName = formField->GetName();
wstring formFieldNameType = getFormFieldType(formField->GetFormFieldType());
```

---

# spire.doc cpp form fields
## get collection of form fields from document
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(inputFile.c_str());

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<FormFieldCollection> formFields = section->GetBody()->GetFormFields();
```

---

# Spire.Doc C++ Get Merge Field Name
## Retrieve names of merge fields from a Word document
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();
document->LoadFromFile(DATAPATH"/MailMerge.doc");

//Get merge field name
std::vector<LPCWSTR_S> fieldNames = document->GetMailMerge()->GetMergeFieldNames();

document->Close();
```

---

# spire.doc cpp fields
## insert address block field
```cpp
//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add address block in the paragraph
intrusive_ptr<Field> field = par->AppendField(L"ADDRESSBLOCK", FieldType::FieldAddressBlock);

//Set field code
field->SetCode(L"ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"");
```

---

# spire.doc cpp advance field
## insert advance field in word document
```cpp
//Open a Word document.
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add advance field
intrusive_ptr<Field> field = par->AppendField(L"Field", FieldType::FieldAdvance);

//Add field code
field->SetCode(L"ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ");

//Update field
document->SetIsUpdateFields(true);
```

---

# Spire.Doc C++ Merge Field
## Insert merge field into a document
```cpp
//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add merge field in the paragraph
intrusive_ptr<MergeField> field = Object::Dynamic_cast<MergeField>(par->AppendField(L"MyFieldName", FieldType::FieldMergeField));
```

---

# spire.doc cpp field
## insert none field into document
```cpp
//Create a Word document
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add a none field
intrusive_ptr<Field> field = par->AppendField(L"", FieldType::FieldNone);
```

---

# Spire.Doc C++ Page Reference Field
## Insert a page reference field into a Word document
```cpp
//Get the first section
intrusive_ptr<Section> section = document->GetLastSection();

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add page ref field
intrusive_ptr<Field> field = par->AppendField(L"pageRef", FieldType::FieldPageRef);

//Set field code
field->SetCode(L"PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT");

//Update field
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp fields
## remove custom property fields from word document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Get custom document properties object.
intrusive_ptr<CustomDocumentProperties> cdp = document->GetCustomDocumentProperties();

//Remove all custom property fields in the document.
for (int i = 0; i < cdp->GetCount();/* i++*/)
{
    cdp->Remove(cdp->GetItem(i)->GetName());
}

document->SetIsUpdateFields(true);
```

---

# spire.doc cpp field
## remove field from document
```cpp
//Get the first field
intrusive_ptr<Field> field = document->GetFields()->GetItem(0);

//Get the paragraph of the field
intrusive_ptr<Paragraph> par = field->GetOwnerParagraph();
//Get the index of the field
int index = par->GetChildObjects()->IndexOf(field);
//Remove if field via index
par->GetChildObjects()->RemoveAt(index);
```

---

# spire.doc cpp field replacement
## replace text with merge field
```cpp
//Open a Word document
intrusive_ptr<Document> document = new Document();

//Find the text that will be replaced
intrusive_ptr<TextSelection> ts = document->FindString(L"Test", true, true);

intrusive_ptr<TextRange> tr = ts->GetAsOneRange();

//Get the paragraph
intrusive_ptr<Paragraph> par = tr->GetOwnerParagraph();

//Get the index of the text in the paragraph
int index = par->GetChildObjects()->IndexOf(tr);

//Create a new field
intrusive_ptr<MergeField> field = new MergeField(document);
field->SetFieldName(L"MergeField");

//Insert field at specific position
par->GetChildObjects()->Insert(index, field);

//Remove the text
par->GetChildObjects()->Remove(tr);

document->Close();
```

---

# spire.doc cpp field culture
## set culture for date field
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();

//Create a section
intrusive_ptr<Section> section = document->AddSection();

//Add paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Add textRnage for paragraph
paragraph->AppendText(L"Add Date Field: ");

//Add date field1
intrusive_ptr<Field> field1 = Object::Dynamic_cast<Field>(paragraph->AppendField(L"Date1", FieldType::FieldDate));
wstring codeString = L"DATE  \\@";
codeString += L"\"yyyy\\MM\\dd\"";
field1->SetCode(codeString.c_str());

//Add new paragraph
intrusive_ptr<Paragraph> newParagraph = section->AddParagraph();

//Add textRnage for paragraph
newParagraph->AppendText(L"Add Date Field with setting French Culture: ");

//Add date field2
intrusive_ptr<Field> field2 = newParagraph->AppendField(L"\"\\@\"dd MMMM yyyy", FieldType::FieldDate);

//Setting Field with setting French Culture
field2->GetCharacterFormat()->SetLocaleIdASCII(1036);

//Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp field locale
## set locale for field in document
```cpp
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

intrusive_ptr<Paragraph> par = section->AddParagraph();

//Add a date field
intrusive_ptr<Field> field = par->AppendField(L"DocDate", FieldType::FieldDate);

//Set the LocaleId for the textrange
(Object::Dynamic_cast<TextRange>(field->GetOwnerParagraph()->GetChildObjects()->GetItem(0)))->GetCharacterFormat()->SetLocaleIdASCII(1049);

field->SetFieldText(L"2019-10-10");
//Update field
document->SetIsUpdateFields(true);
```

---

# tinyxml2 xml parser
## lightweight xml parsing and processing library
```cpp
namespace tinyxml2
{

struct Entity {
    const char* pattern;
    int length;
    char value;
};

static const int NUM_ENTITIES = 5;
static const Entity entities[NUM_ENTITIES] = {
    { "quot", 4,	DOUBLE_QUOTE },
    { "amp", 3,		'&'  },
    { "apos", 4,	SINGLE_QUOTE },
    { "lt",	2, 		'<'	 },
    { "gt",	2,		'>'	 }
};

// --------- XMLUtil ----------- //
const char* XMLUtil::ReadBOM( const char* p, bool* bom )
{
    TIXMLASSERT( p );
    TIXMLASSERT( bom );
    *bom = false;
    const unsigned char* pu = reinterpret_cast<const unsigned char*>(p);
    // Check for BOM:
    if (    *(pu+0) == TIXML_UTF_LEAD_0
            && *(pu+1) == TIXML_UTF_LEAD_1
            && *(pu+2) == TIXML_UTF_LEAD_2 ) {
        *bom = true;
        p += 3;
    }
    TIXMLASSERT( p );
    return p;
}

void XMLUtil::ConvertUTF32ToUTF8( unsigned long input, char* output, int* length )
{
    const unsigned long BYTE_MASK = 0xBF;
    const unsigned long BYTE_MARK = 0x80;
    const unsigned long FIRST_BYTE_MARK[7] = { 0x00, 0x00, 0xC0, 0xE0, 0xF0, 0xF8, 0xFC };

    if (input < 0x80) {
        *length = 1;
    }
    else if ( input < 0x800 ) {
        *length = 2;
    }
    else if ( input < 0x10000 ) {
        *length = 3;
    }
    else if ( input < 0x200000 ) {
        *length = 4;
    }
    else {
        *length = 0;    // This code won't convert this correctly anyway.
        return;
    }

    output += *length;

    switch (*length) {
        case 4:
            --output;
            *output = static_cast<char>((input | BYTE_MARK) & BYTE_MASK);
            input >>= 6;
        case 3:
            --output;
            *output = static_cast<char>((input | BYTE_MARK) & BYTE_MASK);
            input >>= 6;
        case 2:
            --output;
            *output = static_cast<char>((input | BYTE_MARK) & BYTE_MASK);
            input >>= 6;
        case 1:
            --output;
            *output = static_cast<char>(input | FIRST_BYTE_MARK[*length]);
            break;
        default:
            TIXMLASSERT( false );
    }
}

// --------- XMLNode ----------- //
XMLNode::XMLNode( XMLDocument* doc ) :
    _document( doc ),
    _parent( 0 ),
    _value(),
    _parseLineNum( 0 ),
    _firstChild( 0 ), _lastChild( 0 ),
    _prev( 0 ), _next( 0 ),
	_userData( 0 ),
    _memPool( 0 )
{
}

XMLNode::~XMLNode()
{
    DeleteChildren();
    if ( _parent ) {
        _parent->Unlink( this );
    }
}

const char* XMLNode::Value() const
{
    if ( this->ToDocument() )
        return 0;
    return _value.GetStr();
}

void XMLNode::SetValue( const char* str, bool staticMem )
{
    if ( staticMem ) {
        _value.SetInternedStr( str );
    }
    else {
        _value.SetStr( str );
    }
}

void XMLNode::DeleteChildren()
{
    while( _firstChild ) {
        TIXMLASSERT( _lastChild );
        DeleteChild( _firstChild );
    }
    _firstChild = _lastChild = 0;
}

void XMLNode::Unlink( XMLNode* child )
{
    TIXMLASSERT( child );
    TIXMLASSERT( child->_document == _document );
    TIXMLASSERT( child->_parent == this );
    if ( child == _firstChild ) {
        _firstChild = _firstChild->_next;
    }
    if ( child == _lastChild ) {
        _lastChild = _lastChild->_prev;
    }

    if ( child->_prev ) {
        child->_prev->_next = child->_next;
    }
    if ( child->_next ) {
        child->_next->_prev = child->_prev;
    }
	child->_next = 0;
	child->_prev = 0;
	child->_parent = 0;
}

void XMLNode::DeleteChild( XMLNode* node )
{
    TIXMLASSERT( node );
    TIXMLASSERT( node->_document == _document );
    TIXMLASSERT( node->_parent == this );
    Unlink( node );
	TIXMLASSERT(node->_prev == 0);
	TIXMLASSERT(node->_next == 0);
	TIXMLASSERT(node->_parent == 0);
    DeleteNode( node );
}

XMLNode* XMLNode::InsertEndChild( XMLNode* addThis )
{
    TIXMLASSERT( addThis );
    if ( addThis->_document != _document ) {
        TIXMLASSERT( false );
        return 0;
    }
    InsertChildPreamble( addThis );

    if ( _lastChild ) {
        TIXMLASSERT( _firstChild );
        TIXMLASSERT( _lastChild->_next == 0 );
        _lastChild->_next = addThis;
        addThis->_prev = _lastChild;
        _lastChild = addThis;

        addThis->_next = 0;
    }
    else {
        TIXMLASSERT( _firstChild == 0 );
        _firstChild = _lastChild = addThis;

        addThis->_prev = 0;
        addThis->_next = 0;
    }
    addThis->_parent = this;
    return addThis;
}

char* XMLNode::ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr )
{
	XMLDocument::DepthTracker tracker(_document);
	if (_document->Error())
		return 0;

	while( p && *p ) {
        XMLNode* node = 0;

        p = _document->Identify( p, &node );
        TIXMLASSERT( p );
        if ( node == 0 ) {
            break;
        }

        const int initialLineNum = node->_parseLineNum;

        StrPair endTag;
        p = node->ParseDeep( p, &endTag, curLineNumPtr );
        if ( !p ) {
            _document->DeleteNode( node );
            if ( !_document->Error() ) {
                _document->SetError( XML_ERROR_PARSING, initialLineNum, 0);
            }
            break;
        }

        XMLElement* ele = node->ToElement();
        if ( ele ) {
            if ( ele->ClosingType() == XMLElement::CLOSING ) {
                if ( parentEndTag ) {
                    ele->_value.TransferTo( parentEndTag );
                }
                node->_memPool->SetTracked();
                DeleteNode( node );
                return p;
            }

            bool mismatch = false;
            if ( endTag.Empty() ) {
                if ( ele->ClosingType() == XMLElement::OPEN ) {
                    mismatch = true;
                }
            }
            else {
                if ( ele->ClosingType() != XMLElement::OPEN ) {
                    mismatch = true;
                }
                else if ( !XMLUtil::StringEqual( endTag.GetStr(), ele->Name() ) ) {
                    mismatch = true;
                }
            }
            if ( mismatch ) {
                _document->SetError( XML_ERROR_MISMATCHED_ELEMENT, initialLineNum, "XMLElement name=%s", ele->Name());
                _document->DeleteNode( node );
                break;
            }
        }
        InsertEndChild( node );
    }
    return 0;
}

// --------- XMLDocument ----------- //
XMLDocument::XMLDocument( bool processEntities, Whitespace whitespaceMode ) :
    XMLNode( 0 ),
    _writeBOM( false ),
    _processEntities( processEntities ),
    _errorID(XML_SUCCESS),
    _whitespaceMode( whitespaceMode ),
    _errorStr(),
    _errorLineNum( 0 ),
    _charBuffer( 0 ),
    _parseCurLineNum( 0 ),
	_parsingDepth(0),
    _unlinked(),
    _elementPool(),
    _attributePool(),
    _textPool(),
    _commentPool()
{
    _document = this;
}

XMLDocument::~XMLDocument()
{
    Clear();
}

void XMLDocument::Clear()
{
    DeleteChildren();
	while( _unlinked.Size()) {
		DeleteNode(_unlinked[0]);
	}

    ClearError();

    delete [] _charBuffer;
    _charBuffer = 0;
	_parsingDepth = 0;
}

XMLElement* XMLDocument::NewElement( const char* name )
{
    XMLElement* ele = CreateUnlinkedNode<XMLElement>( _elementPool );
    ele->SetName( name );
    return ele;
}

XMLComment* XMLDocument::NewComment( const char* str )
{
    XMLComment* comment = CreateUnlinkedNode<XMLComment>( _commentPool );
    comment->SetValue( str );
    return comment;
}

XMLText* XMLDocument::NewText( const char* str )
{
    XMLText* text = CreateUnlinkedNode<XMLText>( _textPool );
    text->SetValue( str );
    return text;
}

XMLDeclaration* XMLDocument::NewDeclaration( const char* str )
{
    XMLDeclaration* dec = CreateUnlinkedNode<XMLDeclaration>( _commentPool );
    dec->SetValue( str ? str : "xml version=\"1.0\" encoding=\"UTF-8\"" );
    return dec;
}

XMLUnknown* XMLDocument::NewUnknown( const char* str )
{
    XMLUnknown* unk = CreateUnlinkedNode<XMLUnknown>( _commentPool );
    unk->SetValue( str );
    return unk;
}

XMLError XMLDocument::Parse( const char* xml, size_t nBytes )
{
    Clear();

    if ( nBytes == 0 || !xml || !*xml ) {
        SetError( XML_ERROR_EMPTY_DOCUMENT, 0, 0 );
        return _errorID;
    }
    if ( nBytes == static_cast<size_t>(-1) ) {
        nBytes = strlen( xml );
    }
    TIXMLASSERT( _charBuffer == 0 );
    _charBuffer = new char[ nBytes+1 ];
    memcpy( _charBuffer, xml, nBytes );
    _charBuffer[nBytes] = 0;

    Parse();
    if ( Error() ) {
        DeleteChildren();
        _elementPool.Clear();
        _attributePool.Clear();
        _textPool.Clear();
        _commentPool.Clear();
    }
    return _errorID;
}

void XMLDocument::Parse()
{
    TIXMLASSERT( NoChildren() );
    TIXMLASSERT( _charBuffer );
    _parseCurLineNum = 1;
    _parseLineNum = 1;
    char* p = _charBuffer;
    p = XMLUtil::SkipWhiteSpace( p, &_parseCurLineNum );
    p = const_cast<char*>( XMLUtil::ReadBOM( p, &_writeBOM ) );
    if ( !*p ) {
        SetError( XML_ERROR_EMPTY_DOCUMENT, 0, 0 );
        return;
    }
    ParseDeep(p, 0, &_parseCurLineNum );
}

// --------- XMLElement ---------- //
XMLElement::XMLElement( XMLDocument* doc ) : XMLNode( doc ),
    _closingType( OPEN ),
    _rootAttribute( 0 )
{
}

XMLElement::~XMLElement()
{
    while( _rootAttribute ) {
        XMLAttribute* next = _rootAttribute->_next;
        DeleteAttribute( _rootAttribute );
        _rootAttribute = next;
    }
}

const XMLAttribute* XMLElement::FindAttribute( const char* name ) const
{
    for( XMLAttribute* a = _rootAttribute; a; a = a->_next ) {
        if ( XMLUtil::StringEqual( a->Name(), name ) ) {
            return a;
        }
    }
    return 0;
}

const char* XMLElement::GetText() const
{
    const XMLNode* node = FirstChild();
    while (node) {
        if (node->ToComment()) {
            node = node->NextSibling();
            continue;
        }
        break;
    }

    if ( node && node->ToText() ) {
        return node->Value();
    }
    return 0;
}

char* XMLElement::ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr )
{
    p = XMLUtil::SkipWhiteSpace( p, curLineNumPtr );

    if ( *p == '/' ) {
        _closingType = CLOSING;
        ++p;
    }

    p = _value.ParseName( p );
    if ( _value.Empty() ) {
        return 0;
    }

    p = ParseAttributes( p, curLineNumPtr );
    if ( !p || !*p || _closingType != OPEN ) {
        return p;
    }

    p = XMLNode::ParseDeep( p, parentEndTag, curLineNumPtr );
    return p;
}

// --------- XMLAttribute ---------- //
const char* XMLAttribute::Name() const
{
    return _name.GetStr();
}

const char* XMLAttribute::Value() const
{
    return _value.GetStr();
}

char* XMLAttribute::ParseDeep( char* p, bool processEntities, int* curLineNumPtr )
{
    p = _name.ParseName( p );
    if ( !p || !*p ) {
        return 0;
    }

    p = XMLUtil::SkipWhiteSpace( p, curLineNumPtr );
    if ( *p != '=' ) {
        return 0;
    }

    ++p;
    p = XMLUtil::SkipWhiteSpace( p, curLineNumPtr );
    if ( *p != '\"' && *p != '\'' ) {
        return 0;
    }

    const char endTag[2] = { *p, 0 };
    ++p;

    p = _value.ParseText( p, endTag, processEntities ? StrPair::ATTRIBUTE_VALUE : StrPair::ATTRIBUTE_VALUE_LEAVE_ENTITIES, curLineNumPtr );
    return p;
}

}   // namespace tinyxml2
```

---

# tinyxml2 xml parser
## lightweight XML parser library header

```cpp
namespace tinyxml2
{
class XMLDocument;
class XMLElement;
class XMLAttribute;
class XMLComment;
class XMLText;
class XMLDeclaration;
class XMLUnknown;
class XMLPrinter;

/*
	A class that wraps strings. Normally stores the start and end
	pointers into the XML file itself, and will apply normalization
	and entity translation if actually read. Can also store (and memory
	manage) a traditional char[]
*/
class TINYXML2_LIB StrPair
{
public:
    enum Mode {
        NEEDS_ENTITY_PROCESSING			= 0x01,
        NEEDS_NEWLINE_NORMALIZATION		= 0x02,
        NEEDS_WHITESPACE_COLLAPSING     = 0x04,

        TEXT_ELEMENT		            = NEEDS_ENTITY_PROCESSING | NEEDS_NEWLINE_NORMALIZATION,
        TEXT_ELEMENT_LEAVE_ENTITIES		= NEEDS_NEWLINE_NORMALIZATION,
        ATTRIBUTE_NAME		            = 0,
        ATTRIBUTE_VALUE		            = NEEDS_ENTITY_PROCESSING | NEEDS_NEWLINE_NORMALIZATION,
        ATTRIBUTE_VALUE_LEAVE_ENTITIES  = NEEDS_NEWLINE_NORMALIZATION,
        COMMENT							= NEEDS_NEWLINE_NORMALIZATION
    };

    StrPair() : _flags( 0 ), _start( 0 ), _end( 0 ) {}
    ~StrPair();

    void Set( char* start, char* end, int flags ) {
        Reset();
        _start  = start;
        _end    = end;
        _flags  = flags | NEEDS_FLUSH;
    }

    const char* GetStr();

    bool Empty() const {
        return _start == _end;
    }

    void SetInternedStr( const char* str ) {
        Reset();
        _start = const_cast<char*>(str);
    }

    void SetStr( const char* str, int flags=0 );

    char* ParseText( char* in, const char* endTag, int strFlags, int* curLineNumPtr );
    char* ParseName( char* in );

    void TransferTo( StrPair* other );
	void Reset();
};

/*
	Parent virtual class of a pool for fast allocation
	and deallocation of objects.
*/
class MemPool
{
public:
    MemPool() {}
    virtual ~MemPool() {}

    virtual int ItemSize() const = 0;
    virtual void* Alloc() = 0;
    virtual void Free( void* ) = 0;
    virtual void SetTracked() = 0;
};

/**
	Implements the interface to the "Visitor pattern" (see the Accept() method.)
	If you call the Accept() method, it requires being passed a XMLVisitor
	class to handle callbacks. For nodes that contain other nodes (Document, Element)
	you will get called with a VisitEnter/VisitExit pair. Nodes that are always leafs
	are simply called with Visit().
*/
class TINYXML2_LIB XMLVisitor
{
public:
    virtual ~XMLVisitor() {}

    /// Visit a document.
    virtual bool VisitEnter( const XMLDocument& /*doc*/ )			{
        return true;
    }
    /// Visit a document.
    virtual bool VisitExit( const XMLDocument& /*doc*/ )			{
        return true;
    }

    /// Visit an element.
    virtual bool VisitEnter( const XMLElement& /*element*/, const XMLAttribute* /*firstAttribute*/ )	{
        return true;
    }
    /// Visit an element.
    virtual bool VisitExit( const XMLElement& /*element*/ )			{
        return true;
    }

    /// Visit a declaration.
    virtual bool Visit( const XMLDeclaration& /*declaration*/ )		{
        return true;
    }
    /// Visit a text node.
    virtual bool Visit( const XMLText& /*text*/ )					{
        return true;
    }
    /// Visit a comment node.
    virtual bool Visit( const XMLComment& /*comment*/ )				{
        return true;
    }
    /// Visit an unknown node.
    virtual bool Visit( const XMLUnknown& /*unknown*/ )				{
        return true;
    }
};

// WARNING: must match XMLDocument::_errorNames[]
enum XMLError {
    XML_SUCCESS = 0,
    XML_NO_ATTRIBUTE,
    XML_WRONG_ATTRIBUTE_TYPE,
    XML_ERROR_FILE_NOT_FOUND,
    XML_ERROR_FILE_COULD_NOT_BE_OPENED,
    XML_ERROR_FILE_READ_ERROR,
    XML_ERROR_PARSING_ELEMENT,
    XML_ERROR_PARSING_ATTRIBUTE,
    XML_ERROR_PARSING_TEXT,
    XML_ERROR_PARSING_CDATA,
    XML_ERROR_PARSING_COMMENT,
    XML_ERROR_PARSING_DECLARATION,
    XML_ERROR_PARSING_UNKNOWN,
    XML_ERROR_EMPTY_DOCUMENT,
    XML_ERROR_MISMATCHED_ELEMENT,
    XML_ERROR_PARSING,
    XML_CAN_NOT_CONVERT_TEXT,
    XML_NO_TEXT_NODE,
	XML_ELEMENT_DEPTH_EXCEEDED,

	XML_ERROR_COUNT
};

/*
	Utility functionality.
*/
class TINYXML2_LIB XMLUtil
{
public:
    static const char* SkipWhiteSpace( const char* p, int* curLineNumPtr );
    static char* SkipWhiteSpace( char* const p, int* curLineNumPtr );

    static bool IsWhiteSpace( char p );
    inline static bool IsNameStartChar( unsigned char ch );
    inline static bool IsNameChar( unsigned char ch );
    inline static bool IsPrefixHex( const char* p);
    inline static bool StringEqual( const char* p, const char* q, int nChar=INT_MAX );
    inline static bool IsUTF8Continuation( const char p );

    static const char* ReadBOM( const char* p, bool* hasBOM );
    static const char* GetCharacterRef( const char* p, char* value, int* length );
    static void ConvertUTF32ToUTF8( unsigned long input, char* output, int* length );

    static void ToStr( int v, char* buffer, int bufferSize );
    static void ToStr( unsigned v, char* buffer, int bufferSize );
    static void ToStr( bool v, char* buffer, int bufferSize );
    static void ToStr( float v, char* buffer, int bufferSize );
    static void ToStr( double v, char* buffer, int bufferSize );
	static void ToStr(int64_t v, char* buffer, int bufferSize);
    static void ToStr(uint64_t v, char* buffer, int bufferSize);

    static bool	ToInt( const char* str, int* value );
    static bool ToUnsigned( const char* str, unsigned* value );
    static bool	ToBool( const char* str, bool* value );
    static bool	ToFloat( const char* str, float* value );
    static bool ToDouble( const char* str, double* value );
	static bool ToInt64(const char* str, int64_t* value);
    static bool ToUnsigned64(const char* str, uint64_t* value);
	static void SetBoolSerialization(const char* writeTrue, const char* writeFalse);
};

/** XMLNode is a base class for every object that is in the
	XML Document Object Model (DOM), except XMLAttributes.
	Nodes have siblings, a parent, and children which can
	be navigated. A node is always in a XMLDocument.
	The type of a XMLNode can be queried, and it can
	be cast to its more defined type.
*/
class TINYXML2_LIB XMLNode
{
    friend class XMLDocument;
    friend class XMLElement;
public:
    /// Get the XMLDocument that owns this XMLNode.
    const XMLDocument* GetDocument() const;
    /// Get the XMLDocument that owns this XMLNode.
    XMLDocument* GetDocument();

    /// Safely cast to an Element, or null.
    virtual XMLElement*		ToElement()		{ return 0; }
    /// Safely cast to Text, or null.
    virtual XMLText*		ToText()		{ return 0; }
    /// Safely cast to a Comment, or null.
    virtual XMLComment*		ToComment()		{ return 0; }
    /// Safely cast to a Document, or null.
    virtual XMLDocument*	ToDocument()	{ return 0; }
    /// Safely cast to a Declaration, or null.
    virtual XMLDeclaration*	ToDeclaration()	{ return 0; }
    /// Safely cast to an Unknown, or null.
    virtual XMLUnknown*		ToUnknown()		{ return 0; }

    virtual const XMLElement*		ToElement() const		{ return 0; }
    virtual const XMLText*			ToText() const			{ return 0; }
    virtual const XMLComment*		ToComment() const		{ return 0; }
    virtual const XMLDocument*		ToDocument() const		{ return 0; }
    virtual const XMLDeclaration*	ToDeclaration() const	{ return 0; }
    virtual const XMLUnknown*		ToUnknown() const		{ return 0; }

    const char* Value() const;
    void SetValue( const char* val, bool staticMem=false );
    int GetLineNum() const { return _parseLineNum; }

    /// Get the parent of this node on the DOM.
    const XMLNode*	Parent() const			{ return _parent; }
    XMLNode* Parent()						{ return _parent; }

    /// Returns true if this node has no children.
    bool NoChildren() const					{ return !_firstChild; }

    /// Get the first child node, or null if none exists.
    const XMLNode*  FirstChild() const		{ return _firstChild; }
    XMLNode*		FirstChild()			{ return _firstChild; }

    const XMLElement* FirstChildElement( const char* name = 0 ) const;
    XMLElement* FirstChildElement( const char* name = 0 );

    /// Get the last child node, or null if none exists.
    const XMLNode*	LastChild() const		{ return _lastChild; }
    XMLNode*		LastChild()			{ return _lastChild; }

    const XMLElement* LastChildElement( const char* name = 0 ) const;
    XMLElement* LastChildElement( const char* name = 0 );

    /// Get the previous (left) sibling node of this node.
    const XMLNode*	PreviousSibling() const	{ return _prev; }
    XMLNode*	PreviousSibling()			{ return _prev; }

    const XMLElement*	PreviousSiblingElement( const char* name = 0 ) const;
    XMLElement*	PreviousSiblingElement( const char* name = 0 );

    /// Get the next (right) sibling node of this node.
    const XMLNode*	NextSibling() const		{ return _next; }
    XMLNode*	NextSibling()			{ return _next; }

    const XMLElement*	NextSiblingElement( const char* name = 0 ) const;
    XMLElement*	NextSiblingElement( const char* name = 0 );

    XMLNode* InsertEndChild( XMLNode* addThis );
    XMLNode* LinkEndChild( XMLNode* addThis ) { return InsertEndChild( addThis ); }
    XMLNode* InsertFirstChild( XMLNode* addThis );
    XMLNode* InsertAfterChild( XMLNode* afterThis, XMLNode* addThis );

    void DeleteChildren();
    void DeleteChild( XMLNode* node );

    virtual XMLNode* ShallowClone( XMLDocument* document ) const = 0;
	XMLNode* DeepClone( XMLDocument* target ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const = 0;
    virtual bool Accept( XMLVisitor* visitor ) const = 0;

	void SetUserData(void* userData)	{ _userData = userData; }
	void* GetUserData() const			{ return _userData; }

protected:
    explicit XMLNode( XMLDocument* );
    virtual ~XMLNode();

    virtual char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr);

    XMLDocument*	_document;
    XMLNode*		_parent;
    mutable StrPair	_value;
    int             _parseLineNum;

    XMLNode*		_firstChild;
    XMLNode*		_lastChild;

    XMLNode*		_prev;
    XMLNode*		_next;

	void*			_userdata;
};

/** XML text. */
class TINYXML2_LIB XMLText : public XMLNode
{
    friend class XMLDocument;
public:
    virtual bool Accept( XMLVisitor* visitor ) const;

    virtual XMLText* ToText()			{ return this; }
    virtual const XMLText* ToText() const	{ return this; }

    void SetCData( bool isCData )			{ _isCData = isCData; }
    bool CData() const						{ return _isCData; }

    virtual XMLNode* ShallowClone( XMLDocument* document ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const;

protected:
    explicit XMLText( XMLDocument* doc )	: XMLNode( doc ), _isCData( false )	{}
    virtual ~XMLText()												{}

    char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr );

private:
    bool _isCData;
};

/** An XML Comment. */
class TINYXML2_LIB XMLComment : public XMLNode
{
    friend class XMLDocument;
public:
    virtual XMLComment*	ToComment()					{ return this; }
    virtual const XMLComment* ToComment() const		{ return this; }

    virtual bool Accept( XMLVisitor* visitor ) const;

    virtual XMLNode* ShallowClone( XMLDocument* document ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const;

protected:
    explicit XMLComment( XMLDocument* doc );
    virtual ~XMLComment();

    char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr);
};

/** In correct XML the declaration is the first entry in the file. */
class TINYXML2_LIB XMLDeclaration : public XMLNode
{
    friend class XMLDocument;
public:
    virtual XMLDeclaration*	ToDeclaration()					{ return this; }
    virtual const XMLDeclaration* ToDeclaration() const		{ return this; }

    virtual bool Accept( XMLVisitor* visitor ) const;

    virtual XMLNode* ShallowClone( XMLDocument* document ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const;

protected:
    explicit XMLDeclaration( XMLDocument* doc );
    virtual ~XMLDeclaration();

    char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr );
};

/** Any tag that TinyXML-2 doesn't recognize is saved as an unknown. */
class TINYXML2_LIB XMLUnknown : public XMLNode
{
    friend class XMLDocument;
public:
    virtual XMLUnknown*	ToUnknown()					{ return this; }
    virtual const XMLUnknown* ToUnknown() const		{ return this; }

    virtual bool Accept( XMLVisitor* visitor ) const;

    virtual XMLNode* ShallowClone( XMLDocument* document ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const;

protected:
    explicit XMLUnknown( XMLDocument* doc );
    virtual ~XMLUnknown();

    char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr );
};

/** An attribute is a name-value pair. Elements have an arbitrary
	number of attributes, each with a unique name.
*/
class TINYXML2_LIB XMLAttribute
{
    friend class XMLElement;
public:
    /// The name of the attribute.
    const char* Name() const;

    /// The value of the attribute.
    const char* Value() const;

    int GetLineNum() const { return _parseLineNum; }

    /// The next attribute in the list.
    const XMLAttribute* Next() const { return _next; }

	int	IntValue() const;
	int64_t Int64Value() const;
    uint64_t Unsigned64Value() const;
    unsigned UnsignedValue() const;
    bool	 BoolValue() const;
    double 	 DoubleValue() const;
    float	 FloatValue() const;

    XMLError QueryIntValue( int* value ) const;
    XMLError QueryUnsignedValue( unsigned int* value ) const;
	XMLError QueryInt64Value(int64_t* value) const;
    XMLError QueryUnsigned64Value(uint64_t* value) const;
	XMLError QueryBoolValue( bool* value ) const;
    XMLError QueryDoubleValue( double* value ) const;
    XMLError QueryFloatValue( float* value ) const;

    void SetAttribute( const char* value );
    void SetAttribute( int value );
    void SetAttribute( unsigned value );
	void SetAttribute(int64_t value);
    void SetAttribute(uint64_t value);
    void SetAttribute( bool value );
    void SetAttribute( double value );
    void SetAttribute( float value );

private:
    mutable StrPair _name;
    mutable StrPair _value;
    int             _parseLineNum;
    XMLAttribute*   _next;
    MemPool*        _memPool;
};

/** The element is a container class. It has a value, the element name,
	and can contain other elements, text, comments, and unknowns.
	Elements also contain an arbitrary number of attributes.
*/
class TINYXML2_LIB XMLElement : public XMLNode
{
    friend class XMLDocument;
public:
    const char* Name() const		{ return Value(); }
    void SetName( const char* str, bool staticMem=false )	{ SetValue( str, staticMem ); }

    virtual XMLElement* ToElement()				{ return this; }
    virtual const XMLElement* ToElement() const { return this; }
    virtual bool Accept( XMLVisitor* visitor ) const;

    const char* Attribute( const char* name, const char* value=0 ) const;

	int IntAttribute(const char* name, int defaultValue = 0) const;
    unsigned UnsignedAttribute(const char* name, unsigned defaultValue = 0) const;
	int64_t Int64Attribute(const char* name, int64_t defaultValue = 0) const;
    uint64_t Unsigned64Attribute(const char* name, uint64_t defaultValue = 0) const;
	bool BoolAttribute(const char* name, bool defaultValue = false) const;
    double DoubleAttribute(const char* name, double defaultValue = 0) const;
    float FloatAttribute(const char* name, float defaultValue = 0) const;

    XMLError QueryIntAttribute( const char* name, int* value ) const;
    XMLError QueryUnsignedAttribute( const char* name, unsigned int* value ) const;
	XMLError QueryInt64Attribute(const char* name, int64_t* value) const;
    XMLError QueryUnsigned64Attribute(const char* name, uint64_t* value) const;
	XMLError QueryBoolAttribute( const char* name, bool* value ) const;
    XMLError QueryDoubleAttribute( const char* name, double* value ) const;
    XMLError QueryFloatAttribute( const char* name, float* value ) const;
	XMLError QueryStringAttribute(const char* name, const char** value) const;

	XMLError QueryAttribute( const char* name, int* value ) const;
	XMLError QueryAttribute( const char* name, unsigned int* value ) const;
	XMLError QueryAttribute(const char* name, int64_t* value) const;
    XMLError QueryAttribute(const char* name, uint64_t* value) const;
    XMLError QueryAttribute( const char* name, bool* value ) const;
	XMLError QueryAttribute( const char* name, double* value ) const;
	XMLError QueryAttribute( const char* name, float* value ) const;
	XMLError QueryAttribute(const char* name, const char** value) const;

    void SetAttribute( const char* name, const char* value );
    void SetAttribute( const char* name, int value );
    void SetAttribute( const char* name, unsigned value );
	void SetAttribute(const char* name, int64_t value);
    void SetAttribute(const char* name, uint64_t value);
    void SetAttribute( const char* name, bool value );
    void SetAttribute( const char* name, double value );
    void SetAttribute( const char* name, float value );

    void DeleteAttribute( const char* name );

    const XMLAttribute* FirstAttribute() const { return _rootAttribute; }
    const XMLAttribute* FindAttribute( const char* name ) const;

    const char* GetText() const;
	void SetText( const char* inText );
    void SetText( int value );
    void SetText( unsigned value );
	void SetText(int64_t value);
    void SetText(uint64_t value);
	void SetText( bool value );
    void SetText( double value );
    void SetText( float value );

    XMLError QueryIntText( int* ival ) const;
    XMLError QueryUnsignedText( unsigned* uval ) const;
	XMLError QueryInt64Text(int64_t* uval) const;
	XMLError QueryUnsigned64Text(uint64_t* uval) const;
	XMLError QueryBoolText( bool* bval ) const;
    XMLError QueryDoubleText( double* dval ) const;
    XMLError QueryFloatText( float* fval ) const;

	int IntText(int defaultValue = 0) const;
	unsigned UnsignedText(unsigned defaultValue = 0) const;
	int64_t Int64Text(int64_t defaultValue = 0) const;
    uint64_t Unsigned64Text(uint64_t defaultValue = 0) const;
	bool BoolText(bool defaultValue = false) const;
	double DoubleText(double defaultValue = 0) const;
    float FloatText(float defaultValue = 0) const;

    XMLElement* InsertNewChildElement(const char* name);
    XMLComment* InsertNewComment(const char* comment);
    XMLText* InsertNewText(const char* text);
    XMLDeclaration* InsertNewDeclaration(const char* text);
    XMLUnknown* InsertNewUnknown(const char* text);

    enum ElementClosingType {
        OPEN,		// <foo>
        CLOSED,		// <foo/>
        CLOSING		// </foo>
    };
    ElementClosingType ClosingType() const { return _closingType; }
    virtual XMLNode* ShallowClone( XMLDocument* document ) const;
    virtual bool ShallowEqual( const XMLNode* compare ) const;

protected:
    char* ParseDeep( char* p, StrPair* parentEndTag, int* curLineNumPtr );

private:
    XMLElement( XMLDocument* doc );
    virtual ~XMLElement();
    XMLAttribute* FindOrCreateAttribute( const char* name );
    char* ParseAttributes( char* p, int* curLineNumPtr );
    static void DeleteAttribute( XMLAttribute* attribute );
    XMLAttribute* CreateAttribute();

    enum { BUF_SIZE = 200 };
    ElementClosingType _closingType;
    XMLAttribute* _rootAttribute;
};

enum Whitespace {
    PRESERVE_WHITESPACE,
    COLLAPSE_WHITESPACE
};

/** A Document binds together all the functionality.
	It can be saved, loaded, and printed to the screen.
	All Nodes are connected and allocated to a Document.
	If the Document is deleted, all its Nodes are also deleted.
*/
class TINYXML2_LIB XMLDocument : public XMLNode
{
    friend class XMLElement;
    friend class XMLNode;
    friend class XMLText;
    friend class XMLComment;
    friend class XMLDeclaration;
    friend class XMLUnknown;
public:
    XMLDocument( bool processEntities = true, Whitespace whitespaceMode = PRESERVE_WHITESPACE );
    ~XMLDocument();

    virtual XMLDocument* ToDocument()				{ return this; }
    virtual const XMLDocument* ToDocument() const	{ return this; }

    XMLError Parse( const char* xml, size_t nBytes=static_cast<size_t>(-1) );
    XMLError LoadFile( const char* filename );
    XMLError LoadFile( FILE* );
    XMLError SaveFile( const char* filename, bool compact = false );
    XMLError SaveFile( FILE* fp, bool compact = false );

    bool ProcessEntities() const		{ return _processEntities; }
    Whitespace WhitespaceMode() const	{ return _whitespaceMode; }

    bool HasBOM() const { return _writeBOM; }
    void SetBOM( bool useBOM ) { _writeBOM = useBOM; }

    XMLElement* RootElement()				{ return FirstChildElement(); }
    const XMLElement* RootElement() const	{ return FirstChildElement(); }

    void Print( XMLPrinter* streamer=0 ) const;
    virtual bool Accept( XMLVisitor* visitor ) const;

    XMLElement* NewElement( const char* name );
    XMLComment* NewComment( const char* comment );
    XMLText* NewText( const char* text );
    XMLDeclaration* NewDeclaration( const char* text=0 );
    XMLUnknown* NewUnknown( const char* text );

    void DeleteNode( XMLNode* node );

    void ClearError();

    bool Error() const { return _errorID != XML_SUCCESS; }
    XMLError  ErrorID() const { return _errorID; }
	const char* ErrorName() const;
    static const char* ErrorIDToName(XMLError errorID);
	const char* ErrorStr() const;
    void PrintError() const;
    int ErrorLineNum() const { return _errorLineNum; }

    void Clear();
	void DeepCopy(XMLDocument* target) const;

private:
    bool			_writeBOM;
    bool			_processEntities;
    XMLError		_errorID;
    Whitespace		_whitespaceMode;
    mutable StrPair	_errorStr;
    int             _errorLineNum;
    char*			_charBuffer;
    int				_parseCurLineNum;
	int				_parsingDepth;
	DynArray<XMLNode*, 10> _unlinked;

    MemPoolT< sizeof(XMLElement) >	 _elementPool;
    MemPoolT< sizeof(XMLAttribute) > _attributePool;
    MemPoolT< sizeof(XMLText) >		 _textPool;
    MemPoolT< sizeof(XMLComment) >	 _commentPool;

    void Parse();
    void SetError( XMLError error, int lineNum, const char* format, ... );
	void PushDepth();
	void PopDepth();

    template<class NodeType, int PoolElementSize>
    NodeType* CreateUnlinkedNode( MemPoolT<PoolElementSize>& pool );
};

/**
	A XMLHandle is a class that wraps a node pointer with null checks; this is
	an incredibly useful thing. Note that XMLHandle is not part of the TinyXML-2
	DOM structure. It is a separate utility class.
*/
class TINYXML2_LIB XMLHandle
{
public:
    explicit XMLHandle( XMLNode* node ) : _node( node ) {}
    explicit XMLHandle( XMLNode& node ) : _node( &node ) {}
    XMLHandle( const XMLHandle& ref ) : _node( ref._node ) {}
    XMLHandle& operator=( const XMLHandle& ref )	{ _node = ref._node; return *this; }

    XMLHandle FirstChild() 													{ return XMLHandle( _node ? _node->FirstChild() : 0 ); }
    XMLHandle FirstChildElement( const char* name = 0 )						{ return XMLHandle( _node ? _node->FirstChildElement( name ) : 0 ); }
    XMLHandle LastChild()													{ return XMLHandle( _node ? _node->LastChild() : 0 ); }
    XMLHandle LastChildElement( const char* name = 0 )						{ return XMLHandle( _node ? _node->LastChildElement( name ) : 0 ); }
    XMLHandle PreviousSibling()												{ return XMLHandle( _node ? _node->PreviousSibling() : 0 ); }
    XMLHandle PreviousSiblingElement( const char* name = 0 )				{ return XMLHandle( _node ? _node->PreviousSiblingElement( name ) : 0 ); }
    XMLHandle NextSibling()													{ return XMLHandle( _node ? _node->NextSibling() : 0 ); }
    XMLHandle NextSiblingElement( const char* name = 0 )					{ return XMLHandle( _node ? _node->NextSiblingElement( name ) : 0 ); }

    XMLNode* ToNode()							{ return _node; }
    XMLElement* ToElement() 					{ return ( _node ? _node->ToElement() : 0 ); }
    XMLText* ToText() 							{ return ( _node ? _node->ToText() : 0 ); }
    XMLUnknown* ToUnknown() 					{ return ( _node ? _node->ToUnknown() : 0 ); }
    XMLDeclaration* ToDeclaration() 			{ return ( _node ? _node->ToDeclaration() : 0 ); }

private:
    XMLNode* _node;
};

/**
	A variant of the XMLHandle class for working with const XMLNodes and Documents.
*/
class TINYXML2_LIB XMLConstHandle
{
public:
    explicit XMLConstHandle( const XMLNode* node ) : _node( node ) {}
    explicit XMLConstHandle( const XMLNode& node ) : _node( &node ) {}
    XMLConstHandle( const XMLConstHandle& ref ) : _node( ref._node ) {}
    XMLConstHandle& operator=( const XMLConstHandle& ref )	{ _node = ref._node; return *this; }

    const XMLConstHandle FirstChild() const	{ return XMLConstHandle( _node ? _node->FirstChild() : 0 ); }
    const XMLConstHandle FirstChildElement( const char* name = 0 ) const	{ return XMLConstHandle( _node ? _node->FirstChildElement( name ) : 0 ); }
    const XMLConstHandle LastChild() const	{ return XMLConstHandle( _node ? _node->LastChild() : 0 ); }
    const XMLConstHandle LastChildElement( const char* name = 0 ) const	{ return XMLConstHandle( _node ? _node->LastChildElement( name ) : 0 ); }
    const XMLConstHandle PreviousSibling() const	{ return XMLConstHandle( _node ? _node->PreviousSibling() : 0 ); }
    const XMLConstHandle PreviousSiblingElement( const char* name = 0 ) const	{ return XMLConstHandle( _node ? _node->PreviousSiblingElement( name ) : 0 ); }
    const XMLConstHandle NextSibling() const	{ return XMLConstHandle( _node ? _node->NextSibling() : 0 ); }
    const XMLConstHandle NextSiblingElement( const char* name = 0 ) const	{ return XMLConstHandle( _node ? _node->NextSiblingElement( name ) : 0 ); }

    const XMLNode* ToNode() const				{ return _node; }
    const XMLElement* ToElement() const			{ return ( _node ? _node->ToElement() : 0 ); }
    const XMLText* ToText() const				{ return ( _node ? _node->ToText() : 0 ); }
    const XMLUnknown* ToUnknown() const			{ return ( _node ? _node->ToUnknown() : 0 ); }
    const XMLDeclaration* ToDeclaration() const	{ return ( _node ? _node->ToDeclaration() : 0 ); }

private:
    const XMLNode* _node;
};

/**
	Printing functionality. The XMLPrinter gives you more
	options than the XMLDocument::Print() method.
*/
class TINYXML2_LIB XMLPrinter : public XMLVisitor
{
public:
    XMLPrinter( FILE* file=0, bool compact = false, int depth = 0 );
    virtual ~XMLPrinter()	{}

    void PushHeader( bool writeBOM, bool writeDeclaration );
    void OpenElement( const char* name, bool compactMode=false );
    void PushAttribute( const char* name, const char* value );
    void PushAttribute( const char* name, int value );
    void PushAttribute( const char* name, unsigned value );
	void PushAttribute( const char* name, int64_t value );
	void PushAttribute( const char* name, uint64_t value );
	void PushAttribute( const char* name, bool value );
    void PushAttribute( const char* name, double value );
    virtual void CloseElement( bool compactMode=false );

    void PushText( const char* text, bool cdata=false );
    void PushText( int value );
    void PushText( unsigned value );
	void PushText( int64_t value );
	void PushText( uint64_t value );
	void PushText( bool value );
    void PushText( float value );
    void PushText( double value );

    void PushComment( const char* comment );
    void PushDeclaration( const char* value );
    void PushUnknown( const char* value );

    virtual bool VisitEnter( const XMLDocument& /*doc*/ );
    virtual bool VisitExit( const XMLDocument& /*doc*/ )			{ return true; }

    virtual bool VisitEnter( const XMLElement& element, const XMLAttribute* attribute );
    virtual bool VisitExit( const XMLElement& element );

    virtual bool Visit( const XMLText& text );
    virtual bool Visit( const XMLComment& comment );
    virtual bool Visit( const XMLDeclaration& declaration );
    virtual bool Visit( const XMLUnknown& unknown );

    const char* CStr() const { return _buffer.Mem(); }
    int CStrSize() const { return _buffer.Size(); }
    void ClearBuffer( bool resetToFirstElement = true ) {
        _buffer.Clear();
        _buffer.Push(0);
		_firstElement = resetToFirstElement;
    }

protected:
	virtual bool CompactMode( const XMLElement& )	{ return _compactMode; }
    virtual void PrintSpace( int depth );
    virtual void Print( const char* format, ... );
    virtual void Write( const char* data, size_t size );
    virtual void Putc( char ch );
    inline void Write(const char* data) { Write(data, strlen(data)); }
    void SealElementIfJustOpened();

private:
    void PrepareForNewNode( bool compactMode );
    void PrintString( const char*, bool restrictedEntitySet );

    bool _firstElement;
    FILE* _fp;
    int _depth;
    int _textDepth;
    bool _processEntities;
	bool _compactMode;
    bool _elementJustOpened;
    DynArray< const char*, 10 > _stack;
    enum { ENTITY_RANGE = 64, BUF_SIZE = 200 };
    bool _entityFlag[ENTITY_RANGE];
    bool _restrictedEntityFlag[ENTITY_RANGE];
    DynArray< char, 20 > _buffer;
};

}	// tinyxml2
```

---

# spire.doc cpp fields
## update fields in word document
```cpp
// Create a Word document
intrusive_ptr<Document> document = new Document();

// Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp toc style
## change table of content style in word document
```cpp
//Create document object
intrusive_ptr<Document> doc = new Document();

//Defind a Toc style
intrusive_ptr<ParagraphStyle> tocStyle = Object::Dynamic_cast<ParagraphStyle>(Style::CreateBuiltinStyle(BuiltinStyle::Toc1, doc));
tocStyle->GetCharacterFormat()->SetFontName(L"Aleo");
tocStyle->GetCharacterFormat()->SetFontSize(15.0f);
tocStyle->GetCharacterFormat()->SetTextColor(Color::GetCadetBlue());
doc->GetStyles()->Add(tocStyle);

//Loop through sections
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    //Loop through content of section
    for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
        //Find the structure document tag
        if (Object::CheckType<StructureDocumentTag>(obj))
        {
            intrusive_ptr<StructureDocumentTag> tag = boost::dynamic_pointer_cast<StructureDocumentTag>(obj);
            //Find the paragraph where the TOC1 locates
            for (int k = 0; k < tag->GetChildObjects()->GetCount(); k++)
            {
                intrusive_ptr<DocumentObject> cObj = tag->GetChildObjects()->GetItem(k);
                if (Object::CheckType<Paragraph>(cObj))
                {
                    intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(cObj);
                    if (wcscmp(para->GetStyleName(), L"TOC1") == 0)
                    {
                        //Apply the new style for TOC1 paragraph
                        para->ApplyStyle(tocStyle->GetName());
                    }
                }
            }
        }
    }
}
```

---

# spire.doc cpp table of content
## change TOC tab style
```cpp
//Loop through sections
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    //Loop through content of section
    for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
        //Find the structure document tag
        if (Object::CheckType<StructureDocumentTag>(obj) )
        {
            intrusive_ptr<StructureDocumentTag> tag = boost::dynamic_pointer_cast<StructureDocumentTag>(obj);
            //Find the paragraph where the TOC1 locates
            for (int k = 0; k < tag->GetChildObjects()->GetCount(); k++)
            {
                intrusive_ptr<DocumentObject> cObj = tag->GetChildObjects()->GetItem(k);
                if (Object::CheckType<Paragraph>(cObj) )
                {
                    intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(cObj);
                    if (wcscmp(para->GetStyleName(), L"TOC2") == 0)
                    {
                        //Set the tab style of paragraph
                        for (int a = 0; a < para->GetFormat()->GetTabs()->GetCount(); a++)
                        {
                            intrusive_ptr<Tab> tab = para->GetFormat()->GetTabs()->GetItem(a);
                            tab->SetPosition(tab->GetPosition() + 20);
                            tab->SetTabLeader(TabLeader::NoLeader);
                        }
                    }
                }
            }
        }
    }
}
```

---

# spire.doc cpp table of contents
## create table of contents with default styles
```cpp
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> section = doc->AddSection();
intrusive_ptr<Paragraph> para = section->AddParagraph();
//Create table of content with default switches(\o "1-3" \h \z)
para->AppendTOC(1, 3);

intrusive_ptr<Paragraph> par = section->AddParagraph();
intrusive_ptr<TextRange> tr = par->AppendText(L"Flowers");
tr->GetCharacterFormat()->SetFontSize(30);
par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

//Create paragraph and set the head level
intrusive_ptr<Paragraph> para1 = section->AddParagraph();
para1->AppendText(L"Ornithogalum");
//Apply the Heading1 style
para1->ApplyStyle(BuiltinStyle::Heading1);

intrusive_ptr<Paragraph> para2 = section->AddParagraph();
para2->AppendText(L"Rosa");
//Apply the Heading2 style
para2->ApplyStyle(BuiltinStyle::Heading2);

intrusive_ptr<Paragraph> para3 = section->AddParagraph();
para3->AppendText(L"Hyacinth");
//Apply the Heading3 style
para3->ApplyStyle(BuiltinStyle::Heading3);

//Update TOC
doc->UpdateTableOfContents();
```

---

# spire.doc cpp table of contents
## customize table of content in word document
```cpp
//Create a document
intrusive_ptr<Document> doc = new Document();
//Add a section
intrusive_ptr<Section> section = doc->AddSection();
//Customize table of contents with switches
intrusive_ptr<TableOfContent> toc = new TableOfContent(doc, L"{\\o \"1-3\" \\n 1-1}");
intrusive_ptr<Paragraph> para = section->AddParagraph();
para->GetItems()->Add(toc);
para->AppendFieldMark(FieldMarkType::FieldSeparator);
para->AppendText(L"TOC");
para->AppendFieldMark(FieldMarkType::FieldEnd);
doc->SetTOC(toc);

//Update TOC
doc->UpdateTableOfContents();
```

---

# Remove Table of Contents
## Remove table of contents from a Word document
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();

//Get the first GetBody() from the first section
intrusive_ptr<Body> body = document->GetSections()->GetItemInSectionCollection(0)->GetBody();

//Remove TOC from first GetBody()
intrusive_ptr<Regex> reg = new Regex(L"TOC\\w+");

for (int i = 0; i < body->GetParagraphs()->GetCount(); i++)
{
    if (reg->IsMatch(body->GetParagraphs()->GetItemInParagraphCollection(i)->GetStyleName()))
    {
        body->GetParagraphs()->RemoveAt(i);
        i--;
    }
}
```

---

# spire.doc cpp textbox
## delete table from textbox
```cpp
//Get the first textbox
intrusive_ptr<TextBox> textbox = doc->GetTextBoxes()->GetItem(0);

//Remove the first table from the textbox
textbox->GetBody()->GetTables()->RemoveAt(0);
```

---

# spire.doc cpp textbox
## extract text from textboxes in a word document
```cpp
void ExtractTextFromTables(intrusive_ptr<Table> table)
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
				// Process the extracted text
				std::wstring text = paragraph->GetText();
			}
		}
	}
}

// Verify whether the document contains a textbox or not
if (document->GetTextBoxes()->GetCount() > 0)
{
    // Traverse the document
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
                        // Extract text from paragraph in TextBox
                        if (objt->GetDocumentObjectType() == DocumentObjectType::Paragraph)
                        {
                            std::wstring tempStr = (Object::Dynamic_cast<Paragraph>(objt))->GetText();
                        }

                        // Extract text from Table in TextBox
                        if (objt->GetDocumentObjectType() == DocumentObjectType::Table)
                        {
                            intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(objt);
                            ExtractTextFromTables(table);
                        }
                    }
                }
            }
        }
    }
}
```

---

# spire.doc cpp textbox
## insert image into textbox
```cpp
//Create a new document
intrusive_ptr<Document> doc = new Document();

intrusive_ptr<Section> section = doc->AddSection();
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Append a textbox to paragraph
intrusive_ptr<TextBox> tb = paragraph->AppendTextBox(220, 220);

//Set the position of the textbox
tb->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
tb->GetFormat()->SetHorizontalPosition(50);
tb->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
tb->GetFormat()->SetVerticalPosition(50);

//Set the fill effect of textbox as picture
tb->GetFormat()->GetFillEfects()->SetType(BackgroundType::Picture);

//Fill the textbox with a picture
tb->GetFormat()->GetFillEfects()->SetPicture(DATAPATH"/Spire.Doc.png");
```

---

# Spire.Doc C++ TextBox Table
## Insert a table into a textbox and set its position and style
```cpp
//Create a new document
intrusive_ptr<Document> doc = new Document();

//Add a section
intrusive_ptr<Section> section = doc->AddSection();

//Add a paragraph to the section
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Add a textbox to the paragraph
intrusive_ptr<TextBox> textbox = paragraph->AppendTextBox(300, 100);

//Set the position of the textbox
textbox->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
textbox->GetFormat()->SetHorizontalPosition(140);
textbox->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
textbox->GetFormat()->SetVerticalPosition(50);

//Add text to the textbox
intrusive_ptr<Paragraph> textboxParagraph = textbox->GetBody()->AddParagraph();
intrusive_ptr<TextRange> textboxRange = textboxParagraph->AppendText(L"Table 1");
textboxRange->GetCharacterFormat()->SetFontName(L"Arial");

//Insert table to the textbox
intrusive_ptr<Table> table = textbox->GetBody()->AddTable(true);

//Specify the number of rows and columns of the table
table->ResetCells(4, 4);

//Apply style to the table
table->ApplyStyle(DefaultTableStyle::TableColorful2);
```

---

# spire.doc cpp textbox
## read table from textbox
```cpp
//Get the first textbox
intrusive_ptr<TextBox> textbox = doc->GetTextBoxes()->GetItem(0);

//Get the first table in the textbox
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(textbox->GetBody()->GetTables()->GetItemInTableCollection(0));

wstring stringBuilder;

//Loop through the paragraphs of the table cells and extract them to a .txt file
for (int i = 0; i < table->GetRows()->GetCount(); i++)
{
    intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
    for (int j = 0; j < row->GetCells()->GetCount(); j++)
    {
        intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(j);
        for (int k = 0; k < cell->GetParagraphs()->GetCount(); k++)
        {
            intrusive_ptr<Paragraph> paragraph = cell->GetParagraphs()->GetItemInParagraphCollection(k);
            stringBuilder.append(paragraph->GetText());
            stringBuilder.append(L"\t");
        }
    }
    stringBuilder.append(L"\n");
}
```

---

# spire.doc cpp textbox
## remove textbox from document
```cpp
intrusive_ptr<Document> doc = new Document();

//Remove the first text box
doc->GetTextBoxes()->RemoveAt(0);

//Clear all the text boxes
//doc->GetTextBoxes()->Clear();
```

---

# spire.doc cpp textbox
## insert and format textboxes in a word document
```cpp
void InsertTextbox(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();

	//Insert and format the first textbox.
	intrusive_ptr<TextBox> textBox1 = paragraph->AppendTextBox(240, 35);
	textBox1->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox1->GetFormat()->SetLineColor(Color::GetGray());
	textBox1->GetFormat()->SetLineStyle(TextBoxLineStyle::Simple);
	textBox1->GetFormat()->SetFillColor(Color::GetDarkSeaGreen());
	intrusive_ptr<Paragraph> para = textBox1->GetBody()->AddParagraph();
	intrusive_ptr<TextRange> txtrg = para->AppendText(L"Textbox 1 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetWhite());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the second textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	intrusive_ptr<TextBox> textBox2 = paragraph->AppendTextBox(240, 35);
	textBox2->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox2->GetFormat()->SetLineColor(Color::GetTomato());
	textBox2->GetFormat()->SetLineStyle(TextBoxLineStyle::ThinThick);
	textBox2->GetFormat()->SetFillColor(Color::GetBlue());
	textBox2->GetFormat()->SetLineDashing(LineDashing::Dot);
	para = textBox2->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 2 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetPink());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the third textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	intrusive_ptr<TextBox> textBox3 = paragraph->AppendTextBox(240, 35);
	textBox3->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox3->GetFormat()->SetLineColor(Color::GetViolet());
	textBox3->GetFormat()->SetLineStyle(TextBoxLineStyle::Triple);
	textBox3->GetFormat()->SetFillColor(Color::GetPink());
	textBox3->GetFormat()->SetLineDashing(LineDashing::DashDotDot);
	para = textBox3->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 3 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Color::GetTomato());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
}

//Create a Word document and a section.
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> section = document->AddSection();

InsertTextbox(section);
```

---

# spire.doc cpp textbox
## format textbox in document
```cpp
//Create a new document
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> sec = doc->AddSection();

//Add a text box and append sample text
intrusive_ptr<TextBox> TB = doc->GetSections()->GetItemInSectionCollection(0)->AddParagraph()->AppendTextBox(310, 90);
intrusive_ptr<Paragraph> para = TB->GetBody()->AddParagraph();
intrusive_ptr<TextRange> TR = para->AppendText(L"Using Spire.Doc, developers will find a simple and effective method to endow their applications with rich MS Word features. ");
TR->GetCharacterFormat()->SetFontName(L"Cambria ");
TR->GetCharacterFormat()->SetFontSize(13);

//Set exact position for the text box
TB->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
TB->GetFormat()->SetHorizontalPosition(120);
TB->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
TB->GetFormat()->SetVerticalPosition(100);

//Set line style for the text box
TB->GetFormat()->SetLineStyle(TextBoxLineStyle::Double);
TB->GetFormat()->SetLineColor(Color::GetCornflowerBlue());
TB->GetFormat()->SetLineDashing(LineDashing::Solid);
TB->GetFormat()->SetLineWidth(5);

//Set internal margin for the text box
TB->GetFormat()->GetInternalMargin()->SetTop(15);
TB->GetFormat()->GetInternalMargin()->SetBottom(10);
TB->GetFormat()->GetInternalMargin()->SetLeft(12);
TB->GetFormat()->GetInternalMargin()->SetRight(10);
```

---

# spire.doc cpp watermark
## insert image watermark to word document
```cpp
void InsertImageWatermark(intrusive_ptr<Document> document)
{
	intrusive_ptr<PictureWatermark> picture = new PictureWatermark();
	picture->SetPicture(DATAPATH"/ImageWatermark.png");
	picture->SetScaling(250);
	picture->SetIsWashout(false);
	document->SetWatermark(picture);
}
```

---

# spire.doc cpp watermark
## remove image watermark from document
```cpp
//Set the watermark as null to remove the text and image watermark.
document->SetWatermark(nullptr);
```

---

# spire.doc cpp watermark
## remove text watermark from document
```cpp
//Create Word document.
intrusive_ptr<Document> document = new Document();

//Load the file from disk.
document->LoadFromFile(inputFile.c_str());

//Set the watermark as null to remove the text and image watermark.
document->SetWatermark(nullptr);
```

---

# spire.doc cpp watermark
## insert text watermark to document
```cpp
void InsertTextWatermark(intrusive_ptr<Section> section)
{
	intrusive_ptr<TextWatermark> txtWatermark = new TextWatermark();
	txtWatermark->SetText(L"E-iceblue");
	txtWatermark->SetFontSize(95);
	txtWatermark->SetColor(Color::GetBlue());
	txtWatermark->SetLayout(WatermarkLayout::Diagonal);
	section->GetDocument()->SetWatermark(txtWatermark);
}
```

---

# spire.doc cpp ole extraction
## extract OLE objects from Word document and save them as files
```cpp
// Traverse through all sections of the word document    
for (int s = 0; s < doc->GetSections()->GetCount(); s++)
{
    intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(s);
    // Traverse through all Child Objects in the GetBody() of each section
    for (int i = 0; i < sec->GetBody()->GetChildObjects()->GetCount(); i++)
    {
        intrusive_ptr<DocumentObject> obj = sec->GetBody()->GetChildObjects()->GetItem(i);
        // find the paragraph
        if (Object::CheckType<Paragraph>(obj))
        {
            intrusive_ptr<Paragraph> par = boost::dynamic_pointer_cast<Paragraph>(obj);
            for (int j = 0; j < par->GetChildObjects()->GetCount(); j++)
            {
                intrusive_ptr<DocumentObject> o = par->GetChildObjects()->GetItem(j);
                // check whether the object is OLE
                if (o->GetDocumentObjectType() == DocumentObjectType::OleObject)
                {
                    intrusive_ptr<DocOleObject> Ole = Object::Dynamic_cast<DocOleObject>(o);
                    std::wstring s = Ole->GetObjectType();
                    std::vector<byte> native_data = Ole->GetNativeData();

                    // Process the OLE object based on its type
                    if (wcscmp(s.c_str(), L"AcroExch.Document.DC") == 0)
                    {
                        // Save PDF OLE object to file
                        std::ofstream pdf_file("ExtractOLE.pdf", std::ios::out | std::ofstream::binary);
                        pdf_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
                        pdf_file.close();
                    }
                    else if (wcscmp(s.c_str(), L"Excel.Sheet.8") == 0)
                    {
                        // Save Excel OLE object to file
                        std::ofstream xls_file("ExtractOLE.xls", std::ios::out | std::ofstream::binary);
                        xls_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
                        xls_file.close();
                    }
                    else if (wcscmp(s.c_str(), L"PowerPoint.Show.12") == 0)
                    {
                        // Save PowerPoint OLE object to file
                        std::ofstream pptx_file("ExtractOLE.pptx", std::ios::out | std::ofstream::binary);
                        pptx_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
                        pptx_file.close();
                    }
                }
            }
        }
    }
}
```

---

# spire.doc cpp ole
## insert OLE object into Word document
```cpp
//create a document
intrusive_ptr<Document> doc = new Document();

//add a section
intrusive_ptr<Section> sec = doc->AddSection();

//add a paragraph
intrusive_ptr<Paragraph> par = sec->AddParagraph();

//load the image
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
picture->LoadImageSpire(imagePath.c_str());

//insert the OLE
intrusive_ptr<DocOleObject> obj = par->AppendOleObject(filePath.c_str(), picture, OleObjectType::ExcelWorksheet);
```

---

# spire.doc cpp ole
## insert OLE object as icon via stream
```cpp
//Create word document
intrusive_ptr<Document> doc = new Document();

//add a section
intrusive_ptr<Section> sec = doc->AddSection();

//add a paragraph
intrusive_ptr<Paragraph> par = sec->AddParagraph();

//ole stream
intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());

//load the image
intrusive_ptr<DocPicture> picture = new DocPicture(doc);
picture->LoadImageSpire(inputFile_I.c_str());

//insert the OLE from stream
intrusive_ptr<DocOleObject> obj = par->AppendOleObject(stream, picture, L"zip");

//display as icon
obj->SetDisplayAsIcon(true);
```

---

# spire.doc c++ checkbox content control
## add checkbox content control to word document
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();

intrusive_ptr<Section> section = document->AddSection();

//Add a paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Create StructureDocumentTagInline for document
intrusive_ptr<StructureDocumentTagInline> sdt = new StructureDocumentTagInline(document);

//Add sdt in paragraph
paragraph->GetChildObjects()->Add(sdt);

//Specify the type
sdt->GetSDTProperties()->SetSDTType(SdtType::CheckBox);

//Set properties for control
intrusive_ptr<SdtCheckBox> scb = new SdtCheckBox();
sdt->GetSDTProperties()->SetControlProperties(scb);

//Add textRange format
intrusive_ptr<TextRange> tr = new TextRange(document);
tr->GetCharacterFormat()->SetFontName(L"MS Gothic");
tr->GetCharacterFormat()->SetFontSize(12);

//Add textRange to StructureDocumentTagInline
sdt->GetChildObjects()->Add(tr);

//Set checkBox as checked
scb->SetChecked(true);
```

---

# spire.doc cpp content controls
## adding various content controls to a word document
```cpp
//Create a new word document.
intrusive_ptr<Document> document = new Document();
intrusive_ptr<Section> section = document->AddSection();
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example shows how to add content controls in a Word document.");
paragraph = section->AddParagraph();

//Add Combo Box Content Control.
paragraph = section->AddParagraph();
txtRange = paragraph->AppendText(L"Add Combo Box Content Control:  ");
txtRange->GetCharacterFormat()->SetItalic(true);
intrusive_ptr<StructureDocumentTagInline> sd = new StructureDocumentTagInline(document);
paragraph->GetChildObjects()->Add(sd);
sd->GetSDTProperties()->SetSDTType(SdtType::ComboBox);
intrusive_ptr<SdtComboBox> cb = new SdtComboBox();
intrusive_ptr<SdtListItem> tempVar = new SdtListItem(L"Spire.Doc");
cb->GetListItems()->Add(tempVar);
intrusive_ptr<SdtListItem> tempVar2 = new SdtListItem(L"Spire.XLS");
cb->GetListItems()->Add(tempVar2);
intrusive_ptr<SdtListItem> tempVar3 = new SdtListItem(L"Spire.PDF");
cb->GetListItems()->Add(tempVar3);
sd->GetSDTProperties()->SetControlProperties(cb);
intrusive_ptr<TextRange> rt = new TextRange(document);
rt->SetText(cb->GetListItems()->GetItem(0)->GetDisplayText());
sd->GetSDTContent()->GetChildObjects()->Add(rt);

section->AddParagraph();

//Add Text Content Control.
paragraph = section->AddParagraph();
txtRange = paragraph->AppendText(L"Add Text Content Control:  ");
txtRange->GetCharacterFormat()->SetItalic(true);
sd = new StructureDocumentTagInline(document);
paragraph->GetChildObjects()->Add(sd);
sd->GetSDTProperties()->SetSDTType(SdtType::Text);
intrusive_ptr<SdtText> text = new SdtText(true);
text->SetIsMultiline(true);
sd->GetSDTProperties()->SetControlProperties(text);
rt = new TextRange(document);
rt->SetText(L"Text");
sd->GetSDTContent()->GetChildObjects()->Add(rt);

section->AddParagraph();

//Add Picture Content Control.
paragraph = section->AddParagraph();
txtRange = paragraph->AppendText(L"Add Picture Content Control:  ");
txtRange->GetCharacterFormat()->SetItalic(true);
sd = new StructureDocumentTagInline(document);
paragraph->GetChildObjects()->Add(sd);
sd->GetSDTProperties()->SetSDTType(SdtType::Picture);
intrusive_ptr<DocPicture> pic = new DocPicture(document);
pic->SetWidth(10);
pic->SetHeight(10);
sd->GetSDTContent()->GetChildObjects()->Add(pic);

section->AddParagraph();

//Add Date Picker Content Control.
paragraph = section->AddParagraph();
txtRange = paragraph->AppendText(L"Add Date Picker Content Control:  ");
txtRange->GetCharacterFormat()->SetItalic(true);
sd = new StructureDocumentTagInline(document);
paragraph->GetChildObjects()->Add(sd);
sd->GetSDTProperties()->SetSDTType(SdtType::DatePicker);
intrusive_ptr<SdtDate> date = new SdtDate();
date->SetCalendarType(CalendarType::Default);
date->SetDateFormatSpire(L"yyyy.MM.dd");
date->SetFullDate(DateTime::GetNow());
sd->GetSDTProperties()->SetControlProperties(date);
rt = new TextRange(document);
rt->SetText(L"1990.02.08");
sd->GetSDTContent()->GetChildObjects()->Add(rt);

section->AddParagraph();

//Add Drop-Down List Content Control.
paragraph = section->AddParagraph();
txtRange = paragraph->AppendText(L"Add Drop-Down List Content Control:  ");
txtRange->GetCharacterFormat()->SetItalic(true);
sd = new StructureDocumentTagInline(document);
paragraph->GetChildObjects()->Add(sd);
sd->GetSDTProperties()->SetSDTType(SdtType::DropDownList);
intrusive_ptr<SdtDropDownList> sddl = new SdtDropDownList();
intrusive_ptr<SdtListItem> tempVar4 = new SdtListItem(L"Harry");
sddl->GetListItems()->Add(tempVar4);
intrusive_ptr<SdtListItem> tempVar5 = new SdtListItem(L"Jerry");
sddl->GetListItems()->Add(tempVar5);
sd->GetSDTProperties()->SetControlProperties(sddl);
rt = new TextRange(document);
rt->SetText(sddl->GetListItems()->GetItem(0)->GetDisplayText());
sd->GetSDTContent()->GetChildObjects()->Add(rt);
```

---

# spire.doc c++ richtext content control
## add a rich text content control to a word document
```cpp
//Create a document
intrusive_ptr<Document> document = new Document();

//Add a new section
intrusive_ptr<Section> section = document->AddSection();

//Add a paragraph
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

//Create StructureDocumentTagInline for document
intrusive_ptr<StructureDocumentTagInline> sdt = new StructureDocumentTagInline(document);

//Add sdt in paragraph
paragraph->GetChildObjects()->Add(sdt);

//Specify the type
sdt->GetSDTProperties()->SetSDTType(SdtType::RichText);

//Set displaying text
intrusive_ptr<SdtText> text = new SdtText(true);
text->SetIsMultiline(true);
sdt->GetSDTProperties()->SetControlProperties(text);

//Create a TextRange
intrusive_ptr<TextRange> rt = new TextRange(document);
rt->SetText(L"Welcome to use ");
rt->GetCharacterFormat()->SetTextColor(Color::GetGreen());
sdt->GetSDTContent()->GetChildObjects()->Add(rt);

rt = new TextRange(document);
rt->SetText(L"Spire.Doc");
rt->GetCharacterFormat()->SetTextColor(Color::GetOrangeRed());
sdt->GetSDTContent()->GetChildObjects()->Add(rt);
```

---

# spire.doc cpp structured document tag
## manipulate ComboBox items in Word document
```cpp
//Get the combo box from the file
for (int i = 0; i < doc->GetSections()->GetCount(); i++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(i);
    for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
    {
        intrusive_ptr<DocumentObject> bodyObj = section->GetBody()->GetChildObjects()->GetItem(j);
        if (bodyObj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
        {
            //If SDTType is ComboBox
            if ((Object::Dynamic_cast<StructureDocumentTag>(bodyObj))->GetSDTProperties()->GetSDTType() == SdtType::ComboBox)
            {
                intrusive_ptr<StructureDocumentTag> sdt = Object::Dynamic_cast<StructureDocumentTag>(bodyObj);
                intrusive_ptr<SdtControlProperties> scp = sdt->GetSDTProperties()->GetControlProperties();
            
                intrusive_ptr<SdtComboBox> combo = Object::Convert<SdtComboBox>(scp);
                //Remove the second list item
                combo->GetListItems()->RemoveAt(1);
                //Add a new item
                intrusive_ptr<SdtListItem> item = new SdtListItem(L"D", L"D");
                combo->GetListItems()->Add(item);

                //If the value of list items is "D"
                for (int i = 0; i < combo->GetListItems()->GetCount(); i++)
                {
                    intrusive_ptr<SdtListItem> sdtItem = combo->GetListItems()->GetItem(i);
                    if (wcscmp(sdtItem->GetValue(), L"D") == 0)
                    {
                        //Select the item
                        combo->GetListItems()->SetSelectedValue(sdtItem);
                    }
                }
            }
        }
    }
}
```

---

# spire.doc cpp content control
## get properties of structured document tags in word document
```cpp
class StructureTagsInternal : public ReferenceCounter
{
private:
	std::vector<intrusive_ptr<StructureDocumentTagInline>> m_tagInlines;
	std::vector<intrusive_ptr<StructureDocumentTag>> m_tags;

public:
	std::vector<intrusive_ptr<StructureDocumentTagInline>>& GetTagInlines()
	{
		if (m_tagInlines.empty())
		{
			m_tagInlines = std::vector<intrusive_ptr<StructureDocumentTagInline>>();
		}
		return m_tagInlines;
	}
	void SetTagInlines(std::vector<intrusive_ptr<StructureDocumentTagInline>> value)
	{
		m_tagInlines = value;
	}
	std::vector<intrusive_ptr<StructureDocumentTag>>& GetTags()
	{
		if (m_tags.empty())
		{
			m_tags = std::vector<intrusive_ptr<StructureDocumentTag>>();
		}
		return m_tags;
	}
	void SetTags(std::vector<intrusive_ptr<StructureDocumentTag>> value)
	{
		m_tags = value;
	}
};

intrusive_ptr<StructureTagsInternal> GetAllTags(intrusive_ptr<Document> document) {
	intrusive_ptr<StructureTagsInternal> structureTags = new StructureTagsInternal();
	for (int i = 0; i < document->GetSections()->GetCount(); i++) {
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++) {
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (obj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
			{
				intrusive_ptr<StructureDocumentTag> tagObj = Object::Dynamic_cast<StructureDocumentTag>(obj);
				structureTags->GetTags().push_back(tagObj);
			}
			else if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				for (int k = 0; k < (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetItem(k);
					if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
					{
						structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
					}
				}
			}
			else if (obj->GetDocumentObjectType() == DocumentObjectType::Table)
			{
				for (int k = 0; k < (Object::Dynamic_cast<Table>(obj))->GetRows()->GetCount(); k++)
				{
					intrusive_ptr<TableRow> row = (Object::Dynamic_cast<Table>(obj))->GetRows()->GetItemInRowCollection(k);
					for (int l = 0; l < row->GetCells()->GetCount(); l++)
					{
						intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(l);
						for (int m = 0; m < cell->GetChildObjects()->GetCount(); m++)
						{
							intrusive_ptr<DocumentObject> cellChild = cell->GetChildObjects()->GetItem(m);
							if (cellChild->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
							{
								structureTags->GetTags().push_back(Object::Dynamic_cast<StructureDocumentTag>(cellChild));
							}
							else if (cellChild->GetDocumentObjectType() == DocumentObjectType::Paragraph)
							{
								for (int n = 0; n < (Object::Dynamic_cast<Paragraph>(cellChild))->GetChildObjects()->GetCount(); n++)
								{
									intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(cellChild))->GetChildObjects()->GetItem(n);
									if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
									{
										structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
									}
								}
							}
						}
					}
				}
			}
		}
	}
	return structureTags;
}
wstring GetSDTType(SdtType value)
{
	switch (value)
	{
	case Spire::Doc::SdtType::None:
		return L"None";
		break;
	case Spire::Doc::SdtType::RichText:
		return L"RichText";
		break;
	case Spire::Doc::SdtType::Bibliography:
		return L"Bibliography";
		break;
	case Spire::Doc::SdtType::Citation:
		return L"Citation";
		break;
	case Spire::Doc::SdtType::ComboBox:
		return L"ComboBox";
		break;
	case Spire::Doc::SdtType::DropDownList:
		return L"DropDownList";
		break;
	case Spire::Doc::SdtType::Picture:
		return L"Picture";
		break;
	case Spire::Doc::SdtType::Text:
		return L"Text";
		break;
	case Spire::Doc::SdtType::Equation:
		return L"Equation";
		break;
	case Spire::Doc::SdtType::DatePicker:
		return L"DatePicker";
		break;
	case Spire::Doc::SdtType::BuildingBlockGallery:
		return L"BuildingBlockGallery";
		break;
	case Spire::Doc::SdtType::DocPartObj:
		return L"DocPartObj";
		break;
	case Spire::Doc::SdtType::Group:
		return L"Group";
		break;
	case Spire::Doc::SdtType::CheckBox:
		return L"CheckBox";
		break;
	}
	return L"";
}

//Get all structureTags in the Word document
intrusive_ptr<StructureTagsInternal> structureTags = GetAllTags(doc);
//Get all StructureDocumentTagInline objects
std::vector<intrusive_ptr<StructureDocumentTagInline>> tagInlines = structureTags->GetTagInlines();
//Get properties of all tagInlines
for (size_t i = 0; i < tagInlines.size(); i++) {
	intrusive_ptr<StructureDocumentTagInline> tagInline = tagInlines[i];
	std::wstring alias = tagInline->GetSDTProperties()->GetAlias();
	double id = tagInline->GetSDTProperties()->GetId();
	std::wstring tag = tagInline->GetSDTProperties()->GetTag();
	std::wstring STDType = GetSDTType(tagInline->GetSDTProperties()->GetSDTType());
}

//Get all StructureDocumentTag objects
std::vector<intrusive_ptr<StructureDocumentTag>> tags = structureTags->GetTags();
//Get properties of all tags
for (size_t i = 0; i < tags.size(); i++) {
	intrusive_ptr<StructureDocumentTag> tag = tags[i];
	std::wstring alias = tag->GetSDTProperties()->GetAlias();
	double id = tag->GetSDTProperties()->GetId();
	std::wstring tagStr = tag->GetSDTProperties()->GetTag();
	std::wstring STDType = GetSDTType(tag->GetSDTProperties()->GetSDTType());
}
```

---

# Spire.Doc C++ Structured Document Tag
## Lock content control content in a Word document
```cpp
intrusive_ptr<Document> doc = new Document();
intrusive_ptr<Section> section = doc->AddSection();
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
paragraph->AppendHTML(htmlString.c_str());

//Create StructureDocumentTag for document
intrusive_ptr<StructureDocumentTag> sdt = new StructureDocumentTag(doc);
intrusive_ptr<Section> section2 = doc->AddSection();
section2->GetBody()->GetChildObjects()->Add(sdt);

//Specify the type
sdt->GetSDTProperties()->SetSDTType(SdtType::RichText);

for (int i = 0; i < section->GetBody()->GetChildObjects()->GetCount(); i++)
{
	intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(i);
	if (obj->GetDocumentObjectType() == DocumentObjectType::Table)
	{
		sdt->GetSDTContent()->GetChildObjects()->Add(obj->Clone());
	}
}

//Lock content
sdt->GetSDTProperties()->SetLockSettings(LockSettingsType::ContentLocked);

doc->GetSections()->Remove(section);
```

---

# spire.doc cpp structured document tag
## remove content controls from word document
```cpp
//Loop through sections
for (int s = 0; s < doc->GetSections()->GetCount(); s++)
{
    intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(s);
    for (int i = 0; i < section->GetBody()->GetChildObjects()->GetCount(); i++)
    {
        //Loop through contents in paragraph
        if (Object::CheckType<Paragraph>(section->GetBody()->GetChildObjects()->GetItem(i)) )
        {
            intrusive_ptr<Paragraph> para = boost::dynamic_pointer_cast<Paragraph>(section->GetBody()->GetChildObjects()->GetItem(i));
            for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
            {
                //Find the StructureDocumentTagInline
                if (Object::CheckType<StructureDocumentTagInline>(para->GetChildObjects()->GetItem(j)))
                {
                    intrusive_ptr<StructureDocumentTagInline> sdt = boost::dynamic_pointer_cast<StructureDocumentTagInline>(para->GetChildObjects()->GetItem(j));
                    //Remove the content control from paragraph
                    para->GetChildObjects()->Remove(sdt);
                    j--;
                }
            }
        }
        if (Object::CheckType<StructureDocumentTag>(section->GetBody()->GetChildObjects()->GetItem(i)))
        {
            intrusive_ptr<StructureDocumentTag> sdt = boost::dynamic_pointer_cast<StructureDocumentTag>(section->GetBody()->GetChildObjects()->GetItem(i));
            section->GetBody()->GetChildObjects()->Remove(sdt);
            i--;
        }
    }
}
```

---

# spire.doc cpp checkbox
## update checkbox state in structured document
```cpp
class StructureTagsInternal : public ReferenceCounter
{
private:
	std::vector<intrusive_ptr<StructureDocumentTagInline>> m_tagInlines;
	std::vector<intrusive_ptr<StructureDocumentTag>> m_tags;

public:
	std::vector<intrusive_ptr<StructureDocumentTagInline>>& GetTagInlines()
	{
		if (m_tagInlines.empty())
		{
			m_tagInlines = std::vector<intrusive_ptr<StructureDocumentTagInline>>();
		}
		return m_tagInlines;
	}
	void SetTagInlines(std::vector<intrusive_ptr<StructureDocumentTagInline>> value)
	{
		m_tagInlines = value;
	}
	std::vector<intrusive_ptr<StructureDocumentTag>>& GetTags()
	{
		if (m_tags.empty())
		{
			m_tags = std::vector<intrusive_ptr<StructureDocumentTag>>();
		}
		return m_tags;
	}
	void SetTags(std::vector<intrusive_ptr<StructureDocumentTag>> value)
	{
		m_tags = value;
	}
};

intrusive_ptr<StructureTagsInternal> GetAllTags(intrusive_ptr<Document> document)
{
	//Create StructureTags
	intrusive_ptr<StructureTagsInternal> structureTags = new StructureTagsInternal();

	//Travel document sections
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			//Travel document paragraphs
			if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				for (int k = 0; k < (Object::Dynamic_cast <Paragraph> (obj))->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetItem(k);
					//Get StructureDocumentTagInline
					if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
					{
						structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
					}
				}
			}

		}
	}

	return structureTags;
}

wstring GetStdTypeStr(SdtType value)
{
	wstring ret = L"";
	switch (value)
	{
	case Spire::Doc::SdtType::None:
		ret = L"None";
		break;
	case Spire::Doc::SdtType::RichText:
		ret = L"RichText";
		break;
	case Spire::Doc::SdtType::Bibliography:
		ret = L"Bibliography";
		break;
	case Spire::Doc::SdtType::Citation:
		ret = L"Citation";
		break;
	case Spire::Doc::SdtType::ComboBox:
		ret = L"ComboBox";
		break;
	case Spire::Doc::SdtType::DropDownList:
		ret = L"DropDownList";
		break;
	case Spire::Doc::SdtType::Picture:
		ret = L"Picture";
		break;
	case Spire::Doc::SdtType::Text:
		ret = L"Text";
		break;
	case Spire::Doc::SdtType::Equation:
		ret = L"Equation";
		break;
	case Spire::Doc::SdtType::DatePicker:
		ret = L"DatePicker";
		break;
	case Spire::Doc::SdtType::BuildingBlockGallery:
		ret = L"BuildingBlockGallery";
		break;
	case Spire::Doc::SdtType::DocPartObj:
		ret = L"DocPartObj";
		break;
	case Spire::Doc::SdtType::Group:
		ret = L"Group";
		break;
	case Spire::Doc::SdtType::CheckBox:
		ret = L"CheckBox";
		break;
	default:
		break;
	}
	return ret;
}

//Load the document from disk.
document->LoadFromFile(inputFile.c_str());

//Call StructureTags
intrusive_ptr<StructureTagsInternal> structureTags = GetAllTags(document);

//Create list 
vector<intrusive_ptr<StructureDocumentTagInline>>& tagInlines = structureTags->GetTagInlines();

//Get the controls
for (size_t i = 0; i < tagInlines.size(); i++)
{
	//Get the type
	wstring type = GetStdTypeStr(tagInlines[i]->GetSDTProperties()->GetSDTType());

	//Update the status
	if (type == L"CheckBox")
	{
		intrusive_ptr<SdtCheckBox> scb = Object::Dynamic_cast<SdtCheckBox>(tagInlines[i]->GetSDTProperties()->GetControlProperties());
		if (scb->GetChecked())
		{
			scb->SetChecked(false);
		}
		else
		{
			scb->SetChecked(true);
		}
	}
}

//Save the document.
document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
document->Close();
```

---

# spire.doc cpp endnote
## insert and format endnote in document
```cpp
//add endnote to paragraph
intrusive_ptr<Footnote> endnote = p->AppendFootnote(FootnoteType::Endnote);

//append text to endnote
intrusive_ptr<TextRange> text = endnote->GetTextBody()->AddParagraph()->AppendText(L"Reference: Wikipedia");

//set text format
text->GetCharacterFormat()->SetFontName(L"Impact");
text->GetCharacterFormat()->SetFontSize(14);
text->GetCharacterFormat()->SetTextColor(Color::GetDarkOrange());

//Set marker format of endnote
endnote->GetMarkerCharacterFormat()->SetFontName(L"Calibri");
endnote->GetMarkerCharacterFormat()->SetFontSize(25);
endnote->GetMarkerCharacterFormat()->SetTextColor(Color::GetDarkBlue());
```

---

# Spire.Doc C++ Footnote
## Insert and format footnote in a document
```cpp
intrusive_ptr<TextSelection> selection = document->FindString(L"Spire.Doc", false, true);

intrusive_ptr<TextRange> textRange = selection->GetAsOneRange();
intrusive_ptr<Paragraph> paragraph = textRange->GetOwnerParagraph();
int index = paragraph->GetChildObjects()->IndexOf(textRange);
intrusive_ptr<Footnote> footnote = paragraph->AppendFootnote(FootnoteType::Footnote);
paragraph->GetChildObjects()->Insert(index + 1, footnote);

textRange = footnote->GetTextBody()->AddParagraph()->AppendText(L"Welcome to evaluate Spire.Doc");
textRange->GetCharacterFormat()->SetFontName(L"Arial Black");
textRange->GetCharacterFormat()->SetFontSize(10);
textRange->GetCharacterFormat()->SetTextColor(Color::GetDarkGray());

footnote->GetMarkerCharacterFormat()->SetFontName(L"Calibri");
footnote->GetMarkerCharacterFormat()->SetFontSize(12);
footnote->GetMarkerCharacterFormat()->SetBold(true);
footnote->GetMarkerCharacterFormat()->SetTextColor(Color::GetDarkGreen());
```

---

# spire.doc cpp footnote
## remove footnotes from document
```cpp
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

//Traverse paragraphs in the section and find the footnote
for (int k = 0; k < section->GetParagraphs()->GetCount(); k++)
{
	intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(k);
	int index = -1;
	for (int i = 0, cnt = para->GetChildObjects()->GetCount(); i < cnt; i++)
	{
		intrusive_ptr<ParagraphBase> pBase = Object::Dynamic_cast<ParagraphBase>(para->GetChildObjects()->GetItem(i));
		if (Object::CheckType<Footnote>(pBase))
		{
			index = i;
			break;
		}
	}

	if (index > -1)
	{
		//remove the footnote
		para->GetChildObjects()->RemoveAt(index);
	}
}
```

---

# Spire.Doc C++ Footnote Formatting
## Set footnote position and number format
```cpp
//Get the first section
intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(0);

//Set the number format, restart rule and position for the footnote
sec->GetFootnoteOptions()->SetNumberFormat(FootnoteNumberFormat::UpperCaseLetter);
sec->GetFootnoteOptions()->SetRestartRule(FootnoteRestartRule::RestartPage);
sec->GetFootnoteOptions()->SetPosition(FootnotePosition::PrintAsEndOfSection);
```

---

# spire.doc cpp vba macros
## detect and remove vba macros from document
```cpp
//If the document contains Macros, remove them from the document.
if (document->GetIsContainMacro())
{
    document->ClearMacros();
}
```

---

# spire.doc cpp macros
## load and save Word document with macros
```cpp
using namespace Spire::Doc;

//Create a new document
intrusive_ptr<Document> document = new Document();

//Loading document with macros
document->LoadFromFile(inputFile.c_str(), FileFormat::Docm);

//Save docm file
document->SaveToFile(outputFile.c_str(), FileFormat::Docm);
document->Close();
```

---

# Spire.Doc C++ Caption
## Add caption to pictures in Word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Create a new section
intrusive_ptr<Section> section  = document->AddSection();

//Add the first picture
intrusive_ptr<Paragraph> par1 = section->AddParagraph();
par1->GetFormat()->SetAfterSpacing(10);
intrusive_ptr<DocPicture> pic1 = par1->AppendPicture();
pic1->SetHeight(100);
pic1->SetWidth(120);
//Add caption to the picture
pic1->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

//Add the second picture
intrusive_ptr<Paragraph> par2 = section->AddParagraph();
intrusive_ptr<DocPicture> pic2 = par2->AppendPicture();
pic2->SetHeight(100);
pic2->SetWidth(120);
//Add caption to the picture
pic2->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

//Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp table caption
## add caption to table in word document
```cpp
//Get the first table
intrusive_ptr<Body> body = document->GetSections()->GetItemInSectionCollection(0)->GetBody();
intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(body->GetTables()->GetItemInTableCollection(0));

//Add caption to the table
table->AddCaption(L"Table", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

//Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp caption cross reference
## add captions to pictures and create cross references
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Create a new section
intrusive_ptr<Section> section = document->AddSection();

//Add the first paragraph
intrusive_ptr<Paragraph> firstPara = section->AddParagraph();

//Add the first picture
intrusive_ptr<Paragraph> par1 = section->AddParagraph();
par1->GetFormat()->SetAfterSpacing(10);
intrusive_ptr<DocPicture> pic1 = par1->AppendPicture(imagePath1.c_str());
pic1->SetHeight(120);
pic1->SetWidth(120);
//Add caption to the picture
intrusive_ptr<IParagraph> captionParagraph = pic1->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);
section->AddParagraph();

//Add the second picture
intrusive_ptr<Paragraph> par2 = section->AddParagraph();
intrusive_ptr<DocPicture> pic2 = par2->AppendPicture(imagePath2.c_str());
pic2->SetHeight(120);
pic2->SetWidth(120);
//Add caption to the picture
captionParagraph = pic2->AddCaption(L"Figure", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);
section->AddParagraph();

//Create a bookmark
wstring bookmarkName = L"Figure_2";
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
paragraph->AppendBookmarkStart(bookmarkName.c_str());
paragraph->AppendBookmarkEnd(bookmarkName.c_str());

//Replace bookmark content
intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(document);
navigator->MoveToBookmark(bookmarkName.c_str());
intrusive_ptr<TextBodyPart> part = navigator->GetBookmarkContent();
part->GetBodyItems()->Clear();
part->GetBodyItems()->Add(captionParagraph);
navigator->ReplaceBookmarkContent(part);

//Create cross-reference field to point to bookmark "Figure_2"
intrusive_ptr<Field> field = new Field(document);
field->SetType(FieldType::FieldRef);
field->SetCode(L"REF Figure_2 \\p \\h");
firstPara->GetChildObjects()->Add(field);
intrusive_ptr<FieldMark> fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
firstPara->GetChildObjects()->Add(fieldSeparator);

//Set the display text of the field
intrusive_ptr<TextRange> tr = new TextRange(document);
tr->SetText(L"Figure 2");
firstPara->GetChildObjects()->Add(tr);

intrusive_ptr<FieldMark> fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
firstPara->GetChildObjects()->Add(fieldEnd);

//Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp caption
## set caption with chapter number for images in word document
```cpp
//Get the first section
intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
//Label name
std::wstring name = L"Caption ";
for (int i = 0; i < section->GetBody()->GetParagraphs()->GetCount(); i++)
{
    for (int j = 0; j < section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetCount(); j++)
    {
        if (Object::CheckType<DocPicture>(section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetItem(j)))
        {
            intrusive_ptr<DocPicture> pic1 = boost::dynamic_pointer_cast<DocPicture>(section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetItem(j));
            intrusive_ptr<Body> body = Object::Dynamic_cast<Body>(pic1->GetOwnerParagraph()->GetOwner());
            if (body != nullptr)
            {
                int imageIndex = body->GetChildObjects()->IndexOf(pic1->GetOwnerParagraph());
                //Create a new paragraph
                intrusive_ptr<Paragraph> para = new Paragraph(document);
                //Set label
                para->AppendText(name.c_str());

                //Add caption
                intrusive_ptr<Field> field1 = para->AppendField(L"test", FieldType::FieldStyleRef);
                //Chapter number
                field1->SetCode(L" STYLEREF 1 \\s ");
                //Chapter delimiter
                para->AppendText(L" - ");

                //Add picture sequence number
                intrusive_ptr<SequenceField> field2 = Object::Dynamic_cast<SequenceField>(para->AppendField(name.c_str(), FieldType::FieldSequence));
                field2->SetCaptionName(name.c_str());
                field2->SetNumberFormat(CaptionNumberingFormat::Number);
                body->GetParagraphs()->Insert(imageIndex + 1, para);
            }
        }
    }
}
//Set update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp table caption cross-reference
## create table caption and cross-reference in word document
```cpp
//Create word document
intrusive_ptr<Document> document = new Document();

//Get the first section
intrusive_ptr<Section> section = document->AddSection();

//Create a table
intrusive_ptr<Table> table = section->AddTable(true);
table->ResetCells(2, 3);
//Add caption to the table
intrusive_ptr<IParagraph> captionParagraph = table->AddCaption(L"Table", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

//Create a bookmark
wstring bookmarkName = L"Table_1";
intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
paragraph->AppendBookmarkStart(bookmarkName.c_str());
paragraph->AppendBookmarkEnd(bookmarkName.c_str());

//Replace bookmark content
intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(document);
navigator->MoveToBookmark(bookmarkName.c_str());
intrusive_ptr<TextBodyPart> part = navigator->GetBookmarkContent();
part->GetBodyItems()->Clear();
part->GetBodyItems()->Add(captionParagraph);
navigator->ReplaceBookmarkContent(part);

//Create cross-reference field to point to bookmark "Table_1"
intrusive_ptr<Field> field = new Field(document);
field->SetType(FieldType::FieldRef);
field->SetCode(L"REF Table_1 \\p \\h");

//Insert line breaks
for (int i = 0; i < 3; i++)
{
    paragraph->AppendBreak(BreakType::LineBreak);
}

//Insert field to paragraph
paragraph = section->AddParagraph();
intrusive_ptr<TextRange> range = paragraph->AppendText(L"This is a table caption cross-reference, ");
range->GetCharacterFormat()->SetFontSize(14);
paragraph->GetChildObjects()->Add(field);

//Insert FieldSeparator object
intrusive_ptr<FieldMark> fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
paragraph->GetChildObjects()->Add(fieldSeparator);

//Set display text of the field
intrusive_ptr<TextRange> tr = new TextRange(document);
tr->SetText(L"Table 1");
tr->GetCharacterFormat()->SetFontSize(14);
tr->GetCharacterFormat()->SetTextColor(Color::GetDeepSkyBlue());
paragraph->GetChildObjects()->Add(tr);

//Insert FieldEnd object to mark the end of the field
intrusive_ptr<FieldMark> fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
paragraph->GetChildObjects()->Add(fieldEnd);

//Update fields
document->SetIsUpdateFields(true);
```

---

# spire.doc cpp fixed layout
## extract and analyze document layout information
```cpp
// Create a new instance of Document
intrusive_ptr<Document> document = new Document();

// Create a FixedLayoutDocument object
intrusive_ptr<FixedLayoutDocument> layoutDoc = new FixedLayoutDocument(document);

// Get the first line from the first page
intrusive_ptr<FixedLayoutLine> line = layoutDoc->GetPages()->GetItem(0)->GetColumns()->GetItem(0)->GetLines()->GetItem(0);

// Retrieve the original paragraph associated with the line
intrusive_ptr<Paragraph> para = line->GetParagraph();

// Retrieve all the text that appears on the first page in plain text format
wstring pageText = layoutDoc->GetPages()->GetItem(0)->GetText();

// Loop through each page in the document and count lines
for (int i = 0; i < layoutDoc->GetPages()->GetCount(); i++)
{
    intrusive_ptr<FixedLayoutPage> page = layoutDoc->GetPages()->GetItem(i);
    intrusive_ptr<LayoutCollection> lines = page->GetChildEntities(LayoutElementType::Line, true);
}

// Perform a reverse lookup of layout entities for the first paragraph
intrusive_ptr<Paragraph> para2 = (Object::Dynamic_cast<Section>(document->GetFirstChild()))->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0);
intrusive_ptr<LayoutCollection> paragraphLines = layoutDoc->GetLayoutEntitiesOfNode(para2);
for (int i = 0; i < paragraphLines->GetCount(); i++)
{
    intrusive_ptr<FixedLayoutLine> paragraphLine = Object::Dynamic_cast<FixedLayoutLine>(paragraphLines->GetItem(i));
}
```

---


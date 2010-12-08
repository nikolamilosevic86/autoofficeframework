#include "stdafx.h"
#include "MSWord.h"
#include <windows.h>
#include <iostream>
using namespace std;



int main(){
	char ulaz[30];
	CMSWord* word=new CMSWord();
	cout<<"Test application for OLEOfficeAutomationPr.lib. To see help and commands type 'help' \n This test application supports only commands for MS Word \n";
	while(true){
		cout<<"Type command: ";
	cin>>ulaz;
	//cout<<ulaz;
	if(!strcmp(ulaz,"help"))
	{
		cout<<"open - Opens MS Word document. It is needed to type full path to .doc or .docx file \n";
		cout<<"close - Closes active document \n";
		cout<<"quit - Closes MS Word process \n";
		cout<<"delchar - Deletes one char \n";
		cout<<"notvisible - Setting active document not to be visible \n";
		cout<<"visible - Setting active document to be visible \n";
		cout<<"exit - Exits applications\n";
		cout<<"insertText - Inserts a text into MS Word active document. The text has to be typed into console when requested \n";
		cout<<"find - Finds first word that is same as typed \n";
		cout<<"findall - Same as replace :). Replaces text with other \n";
		cout<<"copy - Copies selected text into clipboard \n";
		cout<<"paste - Pastes text from Clipboard to active document \n";
		cout<<"insertPicture - Inserts a picture. Path of picture is parameter \n";
		cout<<"addComment - Adds a commnet in baloon. Text of comment is parameter \n";
		cout<<"newDoc - Creates new document \n";  
		cout<<"count - Counts all open documents \n";  
		cout<<"activateById - Activates document by Id from open documents in the process. \nAnd many more. See the code for more.\n";

		
	}
	if(!strcmp(ulaz,"open"))
	{
	char dokument[150];
	cout<<"Unesite ime dokumenta sa putanjom:";
	cin>>dokument;
	
	//char* dokument="C:\\Projektni.doc";
	CString str(dokument);
	LPCTSTR lpStr = (LPCTSTR)str;
	word->OpenDocument(lpStr,true);
	}
	if(!strcmp(ulaz,"close")){
		word->CloseActiveDocument(true);
	}
	if(!strcmp(ulaz,"quit")){
		word->Quit();
	}
	if(!strcmp(ulaz,"delchar")){
		word->DeleteChar(false);
	}
	if(!strcmp(ulaz,"notvisible")){
		word->SetVisible(false);
	}
	if(!strcmp(ulaz,"visible")){
		word->SetVisible(true);
	}
	if(!strcmp(ulaz,"exit")){
		exit(1);
	}
	if(!strcmp(ulaz,"insertText")){
		char txt[200];
		cout<<"Unesite tekst koji treba uneti:";
		cin>>txt;
		//char* txt="Nikola Milosevic";
		CString str(txt);
		LPCTSTR lpStr = (LPCTSTR)str;
		word->InserText(lpStr);
	}
	if(!strcmp(ulaz,"find")){
	char txt[200];
	cout<<"Unesite tekst koji treba pronaci:";
	cin>>txt;
	//char* txt="Word";
	CString str(txt);
	LPCTSTR lpStr = (LPCTSTR)str;
	word->FindFirst(lpStr);
	}

	if(!strcmp(ulaz,"insertPicture")){
	char txt[200];
	cout<<"Unesite putanju do slike koju treba dodati:";
	cin>>txt;
	//char* txt="Word";
	CString str(txt);
	LPCTSTR lpStr = (LPCTSTR)str;
	word->InserPicture(lpStr);
	}

	if(!strcmp(ulaz,"addComment")){
	char txt[200];
	cout<<"Unesite tekst koji ce biti dodat u komentar:";
	cin>>txt;
	//char* txt="Word";
	CString str(txt);
	LPCTSTR lpStr = (LPCTSTR)str;
	word->AddComment(lpStr);
	}
	if(!strcmp(ulaz,"newDoc")){
	word->NewDocument(true);
	}
	/*if(!strcmp(ulaz,"replace")){
	char* txt="Word";
	CString str(txt);  
	LPCTSTR lpStr = (LPCTSTR)str;  
	char* rpl="Excel";
	CString rstr(rpl);
	LPCTSTR lprepl=(LPCTSTR)rstr;
    word->Replace(lpStr,lprepl,true);   
  } */
	HRESULT hr;
	//Treba ga prebaciti umesto replace all, koji ne radi kako treba zbog argumenta
	if(!strcmp(ulaz,"findall")){
		bool izlaz=true;  
		char txt[200];
		char rpl[200];
		cout<<"Unesite tekst koji treba zameniti:";

		cin>>txt;
		cout<<"Unesite tekst sa kojim treba zameniti predhodno unet tekst:";
		cin>>rpl;
		do{
		
	//char* txt="Word";
	//char txt[200];
	//cin>>txt;
	CString str(txt);
	LPCTSTR lpStr = (LPCTSTR)str;
	izlaz=word->FindFirstBool(lpStr);
	if(!izlaz)break;
	//char* rpl=" Excel ";
	
	CString rstr(rpl);
	LPCTSTR lprepl=(LPCTSTR)rstr;  
	word->DeleteChar(false);
	word->InserText(lprepl);
		}while(izlaz);  
	}
	if(!strcmp(ulaz,"copy")){
	
	word->Copy();  
	}
	if(!strcmp(ulaz,"paste")){
	
	word->Paste();
	}
	if(!strcmp(ulaz,"count")){
	
		cout<<"\nBroj dokumenata je:"<<word->CountDocuments()<<"\n";
	}  

	if(!strcmp(ulaz,"activateById")){
		int a;
		cout<<"id dokumenta za aktiviranje: ";
		cin>>a;
		word->ActivateDocumentById(a);
	}
	/*if(!strcmp(ulaz,"clipboard")){
	 HGLOBAL      temp_Handle ;    // The variable type is case sensitive
	 char*         temp_ptr ;

	 OpenClipboard(0);         // 0 means no window
	 temp_Handle =  GlobalAlloc (GMEM_MOVEABLE + GMEM_DDESHARE,10000 );
	 // temp_ptr = (char*)GlobalLock(temp_Handle);
	 //memcpy (temp_ptr, lpCmdLine, strlen(lpCmdLine)+1);
	  GlobalUnlock(temp_Handle);
	 EmptyClipboard;
	 temp_ptr=(char*)GetClipboardData(CF_TEXT);
	 CloseClipboard();
	 cout<<temp_ptr;


}*/
	
	}
cout<<"Kraj";
}
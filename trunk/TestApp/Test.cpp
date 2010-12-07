#include "stdafx.h"
#include "MSWord.h"
#include <windows.h>
#include <iostream>
using namespace std;



int main(){
	char ulaz[30];
	CMSWord* word=new CMSWord();
	cout<<"Test program za biblioteku OLEOfficeAutomationPr.lib. Da biste videli komande odkucajte 'help' \n";
	while(true){
		cout<<"Unestite komandu: ";
	cin>>ulaz;
	//cout<<ulaz;
	if(!strcmp(ulaz,"help"))
	{
		cout<<"open - sluzi za otvaranje dokumenata, zahteva da se unese puna putanja do .doc ili .docx file-a \n";
		cout<<"close - sluzi za zatvaranje aktivnog dokumenata \n";
		cout<<"quit - sluzi za zatvaranje word procesa \n";
		cout<<"delchar - sluzi za brisanje karaktera \n";
		cout<<"notvisible - sluzi za sakrivanje dokumenta, odnosno aktivan dokument postaje nevidljiv \n";
		cout<<"visible - sluzi da aktivan dokument postane vidljiv \n";
		cout<<"exit - Zatvara program \n";
		cout<<"insertText - Sluzi za dodavanje teksta, tekst se unosi na zahtev programa \n";
		cout<<"find - Sluzi za trazenje zadatog teksta \n";
		cout<<"findall - Sluzi za trazenje zadatog teksa i njegovu zamenu drugim zadatim tekstom \n";
		cout<<"copy - Sluzi kopiranje selektovanog teksta \n";
		cout<<"paste - Sluzi da se nalepi predhodno kopiran \n";
		cout<<"insertPicture - Sluzi da se doda slika u dokument. Slika je parametar \n";
		cout<<"addComment - Sluzi da se doda komentar u dokument. Sadrzaj komentara je parametar \n";
		cout<<"newDoc - Otvara novi dokument \n";  
		cout<<"count - Broji aktivne dokumente u procesu pokrenutom iz aplikacije \n";  
		cout<<"activateById - Aktivira dokument po indexu u listi dokumenata koji sadrzi proces \n";

		
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
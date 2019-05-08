#include<bits/stdc++.h>
using namespace std;
int main(){
	string l;
	string j;
	cout<<"please enter the path of your sws language file that you want to translate to html file _>>";
	cin>>j;
	freopen(j.c_str(),"r",stdin);
	cout<<"<html><body>";
	while(1){
		string u;
		cin>>u;
		if(u=="en"){
			break;
		}else if(u=="l"){
			string a,b,c,f;
			for(;;){
				string y;
				cin>>y;
				if(y=="end!#")break;
				a+=y;
				
			}
			cin>>b>>c>>f;
			cout<<"<p style=\"font-size: "+f+"px;margin-left:"+b+"px;margin-top:"+c+"px\">"+a+"</p>\n";
			
		}
		
	}
	cout<<"</body></html>";
	for(;;);
	return 0;
}


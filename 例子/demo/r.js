var rgb=function(b,g,r){
	return r*65536+g*256+b ;
}

var init=function(){
	Me.ForeColor=rgb(50,137,199);
	Me.font.size=30;
	Me.AutoRedraw=true;
	Me.layout(0,100,100,0);
}

var z=0;

var run=function(data){
	var zc=data.split(" ");
	if(zc[1]=="B0")
		zc[1]=-1;
	Me.cls();
	Me.printf(zc[1]+zc[2]+"."+zc[3]);
	z++;
	if(z>100)
		z=0;
	API.ExecuteAPI("C:/WINDOWS/system32/gdi32.dll", "LineTo " + Me.hdc + ", " + z + " ,100");
	Me.line1(0,0,10,10,rgb(50,137,199));
	Me.line2(20,20,40,40,rgb(50,137,199));
	Me.dpoint(25,25);
	Me.dpoint(26,26);
	Me.dpoint(27,27);
	Me.line3(50,50,60,60,rgb(50,137,199));
    return data;
}




















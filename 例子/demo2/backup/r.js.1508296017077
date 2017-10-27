var rgb=function(b,g,r){
	return r*65536+g*256+b ;
}

var init=function(){
	Me.CurveGraph.Visible=true;
	Me.CurveGraph.CurveCount = 1;
    //水平网格份数：
    Me.CurveGraph.HorizontalSplits = 10;
    //垂直网格份数：
    Me.CurveGraph.VerticalSplits = 30;
    //垂直最小值
    Me.CurveGraph.MinVertical = -200;
    //垂直最大值
    Me.CurveGraph.MaxVertical = 200;
	Me.CurveGraph.CurveLineColor(1) = rgb(255,0,0);
    Me.CurveGraph.ShowGrid = false;

    //注意：坐标轴文字字体 和 X 轴 时间格式 属性，未演示。下面的其实是默认的！！！
    Me.CurveGraph.AxesFont = Me.Font;
    Me.CurveGraph.xBarNowTimeFormat = "hh:mm:ss";
    Me.CurveGraph.FixLegend("xyz");
}

var szc="";
var csh=0;
var agx,agy,agz,qy,wd,xyz;

var run=function(data){
	szc=data;
	agx=szc.split(";")[0];
	agy=szc.split(";")[1];
	agz=szc.split(";")[2];
	//qy=szc.split(";")[3];
	//wd=szc.split(";")[4];
	xyz=Math.sqrt(agx*agx+agy*agy);
	Me.CurveGraph.AddValue(xyz);
	Me.CurveGraph.DrawGridCurve();
	Me.cls();
	//Me.printf("agx:"+agx,0,Me.CurveGraph.Height+10);
	//Me.printf("agy:"+agy,0,Me.CurveGraph.Height+20);
	//Me.printf("agz:"+agz,0,Me.CurveGraph.Height+30);
	Me.printf("xyz:"+xyz,0,Me.CurveGraph.Height+40);
    return data;
}


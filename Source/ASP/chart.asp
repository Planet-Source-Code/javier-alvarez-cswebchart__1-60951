<%

	Response.CacheControl="no-cache"

  Dim WebChart
  Dim vBytes
  
  Set WebChart = Server.CreateObject("CSWebChart.cChart")

	dim FormatType,cbType,PrimaryColor,AlternateColor,LinesType,bValues,bLines,ChartSize,ChartThickness,Png

	FormatType			= ValEx(Request.QueryString("ft"),1)
	nType						= ValEx(Request.QueryString("t"),1)
	PrimaryColor		= ValEx(Request.QueryString("pc"),&HFF7FFFD4)
	AlternateColor	= ValEx(Request.QueryString("ac"),&HFFF0F8FF)
	LinesType				= ValEx(Request.QueryString("tl"),1)
	bValues  				= ValEx(Request.QueryString("bv"),1)
	bLines					= ValEx(Request.QueryString("bl"),1)
	ChartSize				= ValEx(Request.QueryString("cs"),3)
	ChartThickness	= ValEx(Request.QueryString("ct"),3)
	
	if ChartSize = 0 then ChartSize = 200
	if ChartThickness = 0 then ChartThickness = 4
  
  vBytes = WebChart.GetChart(nType, AlternateColor, PrimaryColor, LinesType, bLines, bValues, ChartThickness, ChartSize, FormatType)

	select case FormatType 
		case 0 
			response.ContentType ="image/bmp"
		case 1 
			response.ContentType ="image/jpg"
		case 2 
			response.ContentType ="image/gif"
		case 3 
			response.ContentType ="image/png"
		case 4 
			response.ContentType ="image/bmp"
	end select
	
	Response.BinaryWrite vBytes

function ValEx(value,iDefault)
	if not IsNumeric(value) then
		value = iDefault
	else
		value = cdbl(value)
	end if
	ValEx=value
end function
%>


<html>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<head>
	<title>WebChart Sample</title>
	<link href="./style.css" rel="stylesheet" type="text/css">
</head>
<%

	dim FormatType,cbType,PrimaryColor,AlternateColor,LinesType,bValues,bLines,ChartSize,ChartThickness,Png

	FormatType			= ValEx(Request.Form("cbFormatType"),1)
	nType						= ValEx(Request.Form("cbType"),1)
	PrimaryColor		= ValEx(Request.Form("cbPrimaryColor"),&HFF7FFFD4)
	AlternateColor	= ValEx(Request.Form("cbAlternateColor"),&HFFF0F8FF)
	LinesType				= ValEx(Request.Form("cbLinesType"),1)
	bValues  				= ValEx(Request.Form("cbBarValues"),1)
	bLines					= ValEx(Request.Form("cbOutlines"),1)
	ChartSize				= ValEx(Request.Form("cbChartSize"),3)
	ChartThickness	= ValEx(Request.Form("cbChartThickness"),3)
%>
<body>
	<table>
		<tr>
			<td width=100>&nbsp;
			<td><h1>Graphic Sample</h1>
		<tr>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;
			<td>
						<form name="fchart" method="POST" action="params.asp">
							<table>
								<tr>
									<td>Format
									<td>
										<select name=cbFormatType>
											<option value=0 >BMP</option>
											<option value=1 >JPG</option>
											<option value=2 >GIF</option>
											<option value=3 selected>PNG</option>
											<option value=4 >TIFF</option>
										</select>
								<tr>
									<td>Type
									<td>
										<select name=cbType>
											<option value=0 <%if nType=0 then Response.Write "selected"%>>Pie</option>
											<option value=1 <%if nType=1 then Response.Write "selected"%>>Bar</option>
										</select>
								<tr>
									<td>Bar Primary Color
									<td>
										<select name=cbPrimaryColor>
											<option value=&HFFF0F8FF selected>AliceBlue</option>
											<option value=&HFFFAEBD7>AntiqueWhite </option>
											<option value=&HFF00FFFF>Aqua </option>
											<option value=&HFF7FFFD4>Aquamarine </option>
											<option value=&HFFF0FFFF>Azure </option>
											<option value=&HFFF5F5DC>Beige </option>
											<option value=&HFFFFE4C4>Bisque </option>
											<option value=&HFF000000>Black </option>
											<option value=&HFFFFEBCD>BlanchedAlmond </option>
											<option value=&HFF0000FF>Blue </option>
											<option value=&HFF8A2BE2>BlueViolet </option>
											<option value=&HFFA52A2A>Brown </option>
											<option value=&HFFDEB887>BurlyWood </option>
											<option value=&HFF5F9EA0>CadetBlue </option>
											<option value=&HFF7FFF00>Chartreuse </option>
											<option value=&HFFD2691E>Chocolate </option>
											<option value=&HFFFF7F50>Coral </option>
											<option value=&HFF6495ED>CornflowerBlue </option>
											<option value=&HFFFFF8DC>Cornsilk </option>
											<option value=&HFFDC143C>Crimson </option>
											<option value=&HFF00FFFF>Cyan </option>
											<option value=&HFF00008B>DarkBlue </option>
											<option value=&HFF008B8B>DarkCyan </option>
											<option value=&HFFB8860B>DarkGoldenrod </option>
											<option value=&HFFA9A9A9>DarkGray </option>
											<option value=&HFF006400>DarkGreen </option>
											<option value=&HFFBDB76B>DarkKhaki </option>
											<option value=&HFF8B008B>DarkMagenta </option>
											<option value=&HFF556B2F>DarkOliveGreen </option>
											<option value=&HFFFF8C00>DarkOrange </option>
											<option value=&HFF9932CC>DarkOrchid </option>
											<option value=&HFF8B0000>DarkRed </option>
											<option value=&HFFE9967A>DarkSalmon </option>
											<option value=&HFF8FBC8B>DarkSeaGreen </option>
											<option value=&HFF483D8B>DarkSlateBlue </option>
											<option value=&HFF2F4F4F>DarkSlateGray </option>
											<option value=&HFF00CED1>DarkTurquoise </option>
											<option value=&HFF9400D3>DarkViolet </option>
											<option value=&HFFFF1493>DeepPink </option>
											<option value=&HFF00BFFF>DeepSkyBlue </option>
											<option value=&HFF696969>DimGray </option>
											<option value=&HFF1E90FF>DodgerBlue </option>
											<option value=&HFFB22222>Firebrick </option>
											<option value=&HFFFFFAF0>FloralWhite </option>
											<option value=&HFF228B22>ForestGreen </option>
											<option value=&HFFFF00FF>Fuchsia </option>
											<option value=&HFFDCDCDC>Gainsboro </option>
											<option value=&HFFF8F8FF>GhostWhite </option>
											<option value=&HFFFFD700>Gold </option>
											<option value=&HFFDAA520>Goldenrod </option>
											<option value=&HFF808080>Gray </option>
											<option value=&HFF008000>Green </option>
											<option value=&HFFADFF2F>GreenYellow </option>
											<option value=&HFFF0FFF0>Honeydew </option>
											<option value=&HFFFF69B4>HotPink </option>
											<option value=&HFFCD5C5C>IndianRed </option>
											<option value=&HFF4B0082>Indigo </option>
											<option value=&HFFFFFFF0>Ivory </option>
											<option value=&HFFF0E68C>Khaki </option>
											<option value=&HFFE6E6FA>Lavender </option>
											<option value=&HFFFFF0F5>LavenderBlush </option>
											<option value=&HFF7CFC00>LawnGreen </option>
											<option value=&HFFFFFACD>LemonChiffon </option>
											<option value=&HFFADD8E6>LightBlue </option>
											<option value=&HFFF08080>LightCoral </option>
											<option value=&HFFE0FFFF>LightCyan </option>
											<option value=&HFFFAFAD2>LightGoldenrodYellow </option>
											<option value=&HFFD3D3D3>LightGray </option>
											<option value=&HFF90EE90>LightGreen </option>
											<option value=&HFFFFB6C1>LightPink </option>
											<option value=&HFFFFA07A>LightSalmon </option>
											<option value=&HFF20B2AA>LightSeaGreen </option>
											<option value=&HFF87CEFA>LightSkyBlue </option>
											<option value=&HFF778899>LightSlateGray </option>
											<option value=&HFFB0C4DE>LightSteelBlue </option>
											<option value=&HFFFFFFE0>LightYellow </option>
											<option value=&HFF00FF00>Lime </option>
											<option value=&HFF32CD32>LimeGreen </option>
											<option value=&HFFFAF0E6>Linen </option>
											<option value=&HFFFF00FF>Magenta </option>
											<option value=&HFF800000>Maroon </option>
											<option value=&HFF66CDAA>MediumAquamarine </option>
											<option value=&HFF0000CD>MediumBlue </option>
											<option value=&HFFBA55D3>MediumOrchid </option>
											<option value=&HFF9370DB>MediumPurple </option>
											<option value=&HFF3CB371>MediumSeaGreen </option>
											<option value=&HFF7B68EE>MediumSlateBlue </option>
											<option value=&HFF00FA9A>MediumSpringGreen </option>
											<option value=&HFF48D1CC>MediumTurquoise </option>
											<option value=&HFFC71585>MediumVioletRed </option>
											<option value=&HFF191970>MidnightBlue </option>
											<option value=&HFFF5FFFA>MintCream </option>
											<option value=&HFFFFE4E1>MistyRose </option>
											<option value=&HFFFFE4B5>Moccasin </option>
											<option value=&HFFFFDEAD>NavajoWhite </option>
											<option value=&HFF000080>Navy </option>
											<option value=&HFFFDF5E6>OldLace </option>
											<option value=&HFF808000>Olive </option>
											<option value=&HFF6B8E23>OliveDrab </option>
											<option value=&HFFFFA500>Orange </option>
											<option value=&HFFFF4500>OrangeRed </option>
											<option value=&HFFDA70D6>Orchid </option>
											<option value=&HFFEEE8AA>PaleGoldenrod </option>
											<option value=&HFF98FB98>PaleGreen </option>
											<option value=&HFFAFEEEE>PaleTurquoise </option>
											<option value=&HFFDB7093>PaleVioletRed </option>
											<option value=&HFFFFEFD5>PapayaWhip </option>
											<option value=&HFFFFDAB9>PeachPuff </option>
											<option value=&HFFCD853F>Peru </option>
											<option value=&HFFFFC0CB>Pink </option>
											<option value=&HFFDDA0DD>Plum </option>
											<option value=&HFFB0E0E6>PowderBlue </option>
											<option value=&HFF800080>Purple </option>
											<option value=&HFFFF0000>Red </option>
											<option value=&HFFBC8F8F>RosyBrown </option>
											<option value=&HFF4169E1>RoyalBlue </option>
											<option value=&HFF8B4513>SaddleBrown </option>
											<option value=&HFFFA8072>Salmon </option>
											<option value=&HFFF4A460>SandyBrown </option>
											<option value=&HFF2E8B57>SeaGreen </option>
											<option value=&HFFFFF5EE>SeaShell </option>
											<option value=&HFFA0522D>Sienna </option>
											<option value=&HFFC0C0C0>Silver </option>
											<option value=&HFF87CEEB>SkyBlue </option>
											<option value=&HFF6A5ACD>SlateBlue </option>
											<option value=&HFF708090>SlateGray </option>
											<option value=&HFFFFFAFA>Snow </option>
											<option value=&HFF00FF7F>SpringGreen </option>
											<option value=&HFF4682B4>SteelBlue </option>
											<option value=&HFFD2B48C>Tan </option>
											<option value=&HFF008080>Teal </option>
											<option value=&HFFD8BFD8>Thistle </option>
											<option value=&HFFFF6347>Tomato </option>
											<option value=&HFFFFFF>Transparent </option>
											<option value=&HFF40E0D0>Turquoise </option>
											<option value=&HFFEE82EE>Violet </option>
											<option value=&HFFF5DEB3>Wheat </option>
											<option value=&HFFFFFFFF>White </option>
											<option value=&HFFF5F5F5>WhiteSmoke </option>
											<option value=&HFFFFFF00>Yellow </option>
											<option value=&HFF9ACD32>YellowGreen </option>
										</select>
								<tr>
									<td>Bar Alternate Color
									<td>
										<select name=cbAlternateColor>
											<option value=&HFFF0F8FF>AliceBlue</option>
											<option value=&HFFFAEBD7>AntiqueWhite </option>
											<option value=&HFF00FFFF>Aqua </option>
											<option value=&HFF7FFFD4 selected>Aquamarine </option>
											<option value=&HFFF0FFFF>Azure </option>
											<option value=&HFFF5F5DC>Beige </option>
											<option value=&HFFFFE4C4>Bisque </option>
											<option value=&HFF000000>Black </option>
											<option value=&HFFFFEBCD>BlanchedAlmond </option>
											<option value=&HFF0000FF>Blue </option>
											<option value=&HFF8A2BE2>BlueViolet </option>
											<option value=&HFFA52A2A>Brown </option>
											<option value=&HFFDEB887>BurlyWood </option>
											<option value=&HFF5F9EA0>CadetBlue </option>
											<option value=&HFF7FFF00>Chartreuse </option>
											<option value=&HFFD2691E>Chocolate </option>
											<option value=&HFFFF7F50>Coral </option>
											<option value=&HFF6495ED>CornflowerBlue </option>
											<option value=&HFFFFF8DC>Cornsilk </option>
											<option value=&HFFDC143C>Crimson </option>
											<option value=&HFF00FFFF>Cyan </option>
											<option value=&HFF00008B>DarkBlue </option>
											<option value=&HFF008B8B>DarkCyan </option>
											<option value=&HFFB8860B>DarkGoldenrod </option>
											<option value=&HFFA9A9A9>DarkGray </option>
											<option value=&HFF006400>DarkGreen </option>
											<option value=&HFFBDB76B>DarkKhaki </option>
											<option value=&HFF8B008B>DarkMagenta </option>
											<option value=&HFF556B2F>DarkOliveGreen </option>
											<option value=&HFFFF8C00>DarkOrange </option>
											<option value=&HFF9932CC>DarkOrchid </option>
											<option value=&HFF8B0000>DarkRed </option>
											<option value=&HFFE9967A>DarkSalmon </option>
											<option value=&HFF8FBC8B>DarkSeaGreen </option>
											<option value=&HFF483D8B>DarkSlateBlue </option>
											<option value=&HFF2F4F4F>DarkSlateGray </option>
											<option value=&HFF00CED1>DarkTurquoise </option>
											<option value=&HFF9400D3>DarkViolet </option>
											<option value=&HFFFF1493>DeepPink </option>
											<option value=&HFF00BFFF>DeepSkyBlue </option>
											<option value=&HFF696969>DimGray </option>
											<option value=&HFF1E90FF>DodgerBlue </option>
											<option value=&HFFB22222>Firebrick </option>
											<option value=&HFFFFFAF0>FloralWhite </option>
											<option value=&HFF228B22>ForestGreen </option>
											<option value=&HFFFF00FF>Fuchsia </option>
											<option value=&HFFDCDCDC>Gainsboro </option>
											<option value=&HFFF8F8FF>GhostWhite </option>
											<option value=&HFFFFD700>Gold </option>
											<option value=&HFFDAA520>Goldenrod </option>
											<option value=&HFF808080>Gray </option>
											<option value=&HFF008000>Green </option>
											<option value=&HFFADFF2F>GreenYellow </option>
											<option value=&HFFF0FFF0>Honeydew </option>
											<option value=&HFFFF69B4>HotPink </option>
											<option value=&HFFCD5C5C>IndianRed </option>
											<option value=&HFF4B0082>Indigo </option>
											<option value=&HFFFFFFF0>Ivory </option>
											<option value=&HFFF0E68C>Khaki </option>
											<option value=&HFFE6E6FA>Lavender </option>
											<option value=&HFFFFF0F5>LavenderBlush </option>
											<option value=&HFF7CFC00>LawnGreen </option>
											<option value=&HFFFFFACD>LemonChiffon </option>
											<option value=&HFFADD8E6>LightBlue </option>
											<option value=&HFFF08080>LightCoral </option>
											<option value=&HFFE0FFFF>LightCyan </option>
											<option value=&HFFFAFAD2>LightGoldenrodYellow </option>
											<option value=&HFFD3D3D3>LightGray </option>
											<option value=&HFF90EE90>LightGreen </option>
											<option value=&HFFFFB6C1>LightPink </option>
											<option value=&HFFFFA07A>LightSalmon </option>
											<option value=&HFF20B2AA>LightSeaGreen </option>
											<option value=&HFF87CEFA>LightSkyBlue </option>
											<option value=&HFF778899>LightSlateGray </option>
											<option value=&HFFB0C4DE>LightSteelBlue </option>
											<option value=&HFFFFFFE0>LightYellow </option>
											<option value=&HFF00FF00>Lime </option>
											<option value=&HFF32CD32>LimeGreen </option>
											<option value=&HFFFAF0E6>Linen </option>
											<option value=&HFFFF00FF>Magenta </option>
											<option value=&HFF800000>Maroon </option>
											<option value=&HFF66CDAA>MediumAquamarine </option>
											<option value=&HFF0000CD>MediumBlue </option>
											<option value=&HFFBA55D3>MediumOrchid </option>
											<option value=&HFF9370DB>MediumPurple </option>
											<option value=&HFF3CB371>MediumSeaGreen </option>
											<option value=&HFF7B68EE>MediumSlateBlue </option>
											<option value=&HFF00FA9A>MediumSpringGreen </option>
											<option value=&HFF48D1CC>MediumTurquoise </option>
											<option value=&HFFC71585>MediumVioletRed </option>
											<option value=&HFF191970>MidnightBlue </option>
											<option value=&HFFF5FFFA>MintCream </option>
											<option value=&HFFFFE4E1>MistyRose </option>
											<option value=&HFFFFE4B5>Moccasin </option>
											<option value=&HFFFFDEAD>NavajoWhite </option>
											<option value=&HFF000080>Navy </option>
											<option value=&HFFFDF5E6>OldLace </option>
											<option value=&HFF808000>Olive </option>
											<option value=&HFF6B8E23>OliveDrab </option>
											<option value=&HFFFFA500>Orange </option>
											<option value=&HFFFF4500>OrangeRed </option>
											<option value=&HFFDA70D6>Orchid </option>
											<option value=&HFFEEE8AA>PaleGoldenrod </option>
											<option value=&HFF98FB98>PaleGreen </option>
											<option value=&HFFAFEEEE>PaleTurquoise </option>
											<option value=&HFFDB7093>PaleVioletRed </option>
											<option value=&HFFFFEFD5>PapayaWhip </option>
											<option value=&HFFFFDAB9>PeachPuff </option>
											<option value=&HFFCD853F>Peru </option>
											<option value=&HFFFFC0CB>Pink </option>
											<option value=&HFFDDA0DD>Plum </option>
											<option value=&HFFB0E0E6>PowderBlue </option>
											<option value=&HFF800080>Purple </option>
											<option value=&HFFFF0000>Red </option>
											<option value=&HFFBC8F8F>RosyBrown </option>
											<option value=&HFF4169E1>RoyalBlue </option>
											<option value=&HFF8B4513>SaddleBrown </option>
											<option value=&HFFFA8072>Salmon </option>
											<option value=&HFFF4A460>SandyBrown </option>
											<option value=&HFF2E8B57>SeaGreen </option>
											<option value=&HFFFFF5EE>SeaShell </option>
											<option value=&HFFA0522D>Sienna </option>
											<option value=&HFFC0C0C0>Silver </option>
											<option value=&HFF87CEEB>SkyBlue </option>
											<option value=&HFF6A5ACD>SlateBlue </option>
											<option value=&HFF708090>SlateGray </option>
											<option value=&HFFFFFAFA>Snow </option>
											<option value=&HFF00FF7F>SpringGreen </option>
											<option value=&HFF4682B4>SteelBlue </option>
											<option value=&HFFD2B48C>Tan </option>
											<option value=&HFF008080>Teal </option>
											<option value=&HFFD8BFD8>Thistle </option>
											<option value=&HFFFF6347>Tomato </option>
											<option value=&HFFFFFF>Transparent </option>
											<option value=&HFF40E0D0>Turquoise </option>
											<option value=&HFFEE82EE>Violet </option>
											<option value=&HFFF5DEB3>Wheat </option>
											<option value=&HFFFFFFFF>White </option>
											<option value=&HFFF5F5F5>WhiteSmoke </option>
											<option value=&HFFFFFF00>Yellow </option>
											<option value=&HFF9ACD32>YellowGreen </option>
										</select>
								<tr>
									<td>Bar Grid Lines
									<td>
										<select name=cbLinesType>
											<option  value=0>None</option>
											<option  value=1>Horizontal</option>
											<option  value=2>Numbered</option>
											<option  value=3 selected>Both</option>
										</select>
								<tr>
									<td>Pie Chart Size
									<td>
										<select name=cbChartSize>
											<option  value=50>Smallest</option>
											<option  value=100>Smaller</option>
											<option  value=150>Small</option>
											<option  value=200 selected>Medium</option>
											<option  value=250>Large</option>
											<option  value=350>Larger</option>
										</select>
								<tr>
									<td>Pie Chart Thickness
									<td>
										<select name=cbChartThickness>
											<option value=0>None</option>
											<option value=2>Wafer</option>
											<option value=4 selected>Thin</option>
											<option value=8>Medium</option>
											<option value=16>Thick</option>
											<option value=32>Thickest</option>
										</select>
								<tr>
									<td>Bar Values
									<td>
										<select name=cbBarValues>
											<option value=0>No</option>
											<option value=1 selected>Yes</option>
										</select>
								<tr>
									<td>Show Outlines
									<td>
										<select name=cbOutlines>
											<option value=0>No</option>
											<option value=1 selected>Yes</option>
										</select>
							</table>
							<table>
								<tr>
									<td>
										<input type="submit" value="Crear grafico">
								<tr>
									<td>
						<%

							dim params
							
							params = "ft="  & FormatType
							params = params & "&t="  & nType
							params = params & "&pc=" & PrimaryColor
							params = params & "&ac=" & AlternateColor
							params = params & "&tl=" & LinesType
							params = params & "&bv=" & bValues
							params = params & "&bl=" & bLines
							params = params & "&cs=" & ChartSize
							params = params & "&ct=" & ChartThickness

						function ValEx(value,iDefault)
							if not IsNumeric(value) then
								value = iDefault
							else
								value = cdbl(value)
							end if
							ValEx=value
						end function
						%>	
							<img src="chart.asp?<%=params%>">
							</table>
						</form>
	</table>
</body>
</html>
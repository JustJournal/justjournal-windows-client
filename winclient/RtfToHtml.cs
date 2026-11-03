using System;
using System.Text.RegularExpressions;
using System.Collections;
using System.Windows.Forms;

namespace JustJournal
{
	/// <summary>
	/// Summary description for RtfToHtml.
	/// </summary>
	public static class RtfToHtml
	{
	    public static string ToCssFontSize( int size )
		{
			switch( size )
			{
				case 1:
				case 2:
				case 3:
				case 4:
				case 5:
					return "xx-small";
				case 6:
				case 7:
					return "x-small";
				case 8:
				case 9:
				case 10:
					return "small";
				case 11:
				case 12:
					return "medium";
				case 13:
				case 14:
					return "large";
				case 15:
				case 16:
				case 17:
				case 18:
					return "x-large";
				default:
					return "xx-large";
				
			}
		}

		public static string ToHtmlFontSize( int size )
		{
				switch( size )
				{
					case 1:
					case 2:
					case 3:
					case 4:
					case 5:
					case 6:
					case 7:
						return "1";
					case 8:
					case 9:
					case 10:
						return "2";
					case 11:
					case 12:
						return "3";
					case 13:
					case 14:
						return "4";
					case 15:
					case 16:
					case 17:
					case 18:
						return "5";
					case 19:
					case 20:
					case 21:
					case 22:
					case 23:
					case 24:
						return "6";
					default:
						return "7";
				}
		}

		public static string ToHexColor( int red, int green, int blue )
		{
			return "#" + System.Convert.ToString( red, 16 ).PadLeft( 2, '0' )
				+ System.Convert.ToString( green, 16 ).PadLeft( 2, '0' )
				+ System.Convert.ToString( blue, 16 ).PadLeft( 2, '0' );
		}

		// Mono's RichTextBox writes \cf1 without a \colortbl, so a missing color means none
		private static string TableColor( Hashtable colorTable, string index )
		{
			object color = colorTable[System.Convert.ToInt32( index )];
			return color == null ? string.Empty : color.ToString();
		}

		public static string Convert( string rtf, string bgcolor )
		{
			try
			{
				#region write RTF string to file on desktop
				//StreamWriter sw = File.CreateText( Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\debug.txt" );
				//sw.Write( rtf );
				//sw.Close();
				#endregion
				// drop ignorable groups like {\*\generator Mono RichTextBox;}, which aren't text,
				// and space out Mono's "\colortbl;" so it reads like RichEdit's "\colortbl ;"
				rtf = Regex.Replace( rtf, @"\{\\\*\\[^{}]*\}", string.Empty ).Replace( @"\colortbl;", @"\colortbl ;" );

				Hashtable fontTable = new Hashtable();
				Hashtable colorTable = new Hashtable();
				string result = string.Empty;
				string tg = string.Empty;
				string sz = "2";
				string ft = "Arial Unicode MS";
				bool tag = false;
				int brace = 0;
				bool rdfonts = false;
				int ftlevel = 0;
				string deffont = string.Empty;
				bool rdcolor = false;
				//					bool utext = false;
				

				int cur = 0;

				while( cur < rtf.Length )
				{
					if( !rtf[cur].Equals( '\n' ) ) //skip linefeeds
					{
						if( ( rtf[cur].Equals( ' ' ) || rtf[cur].Equals( '\r' ) || rtf[cur].Equals( '\\' ) ) && tag || rdcolor )
						{
							tag = false;
							//								if( utext && !tg.StartsWith( "'" ) )
							//								{
							//									byte[] ar = (byte[])al.ToArray( typeof( byte ) );
							//									Decoder dec = Encoding.UTF8.GetDecoder();
							//									int numchars = dec.GetCharCount( ar, 0, ar.Length );
							//									char[] chars = new char[numchars];
							//									dec.GetChars( ar, 0, ar.Length, chars, 0 );
							//									result += new string( chars );
							//									utext = false;
							//								}
							if( tg.StartsWith( "deff" ) )
							{
								deffont = tg.Substring( 3, tg.Length - 3 );
							}
							else if( tg.StartsWith( "f" ) ) //could be font or fontsize
							{
								if( tg.Equals( "fonttbl" ) )
								{
									rdfonts = true;
									ftlevel = brace - 1;
								}
								else if( rdfonts ) 
								{
									string ftname = string.Empty;

									while( !rtf[cur].Equals( ' ' ) )
										cur++;

									cur++;

									while( !(rtf[cur].Equals( ';' )) )
										ftname += rtf[cur++];

									fontTable.Add( tg, ftname );

									if( tg.Equals( deffont ) )
										ft = fontTable[deffont].ToString();
								}
								else
								{
									if( TagCounts.Font ) 
									{
										result += "</font>";
										TagCounts.Font = false;
									}
									if( tg.StartsWith( "fs" ) )
									{
										sz = ToHtmlFontSize( System.Convert.ToInt32( tg.Substring( 2 ) ) / 2 );
									}
									else if( fontTable.ContainsKey( tg ) )
									{
										ft = fontTable[tg].ToString();
									}
									else if( fontTable.ContainsKey( deffont ) )
									{
										ft = fontTable[deffont].ToString();
									}
									TagCounts.Font = true;
									result += "<font face=\"" + ft + "\" size=\"" + sz + "\">";
								}
							}
							else if( rdcolor )
							{
								string strcolor = string.Empty;
								int colorindex = 0;
								while( !(rtf[cur].Equals( '}' )) )
								{
									while( !(rtf[cur].Equals( ';' )) )
									{
										strcolor += rtf[cur++];
									}
									if( strcolor.Length > 0 )
									{
										int red = 0, green = 0, blue = 0;
										strcolor = strcolor.TrimStart( '\\' );
										string[] colors = strcolor.Split( '\\' );
										foreach( string s in colors )
										{
											if( s.StartsWith( "red" ) )
											{
												red = System.Convert.ToInt32( s.Substring( 3 ) );
											}
											else if( s.StartsWith( "green" ) )
											{
												green = System.Convert.ToInt32( s.Substring( 5 ) );
											}
											else if( s.StartsWith( "blue" ) )
											{
												blue = System.Convert.ToInt32( s.Substring( 4 ) );
											}
										}
										colorTable.Add( colorindex++, ToHexColor( red, green, blue ) );
									}
									else
									{
										colorTable.Add( colorindex++, string.Empty );
									}
									cur++;
								}
								rdcolor = false;
							}
							else if( tg.StartsWith( "highlight" ) )
							{
								string color = TableColor( colorTable, tg.Substring( 9 ) );
								if( color.Length > 0 )
								{
									result += "<span style=\"background-color:" + color + ";\">";
									TagCounts.Span = true;
								}
							}
							else if( tg.StartsWith( "cf" ) )
							{
								string color = TableColor( colorTable, tg.Substring( 2 ) );
								if( color.Length > 0 )
								{
									result += "<span style=\"color:" + color + ";\">";
									TagCounts.Span = true;
								}
							}
								//								else if( tg.StartsWith( "'" ) )
								//								{
								//									utext = true;
								//									al.Add( System.Convert.ToByte( tg.Substring( 1 ), 16 ) );
								//								}
							else
							{
								switch( tg )
								{
									case "par":
										if( TagCounts.Span )
										{
											result += "</span>";
											TagCounts.Span = false;
										}
										result += "<br />"; // newline/paragraph
										break;
									case "ul":
										result += "<u>";
										break;
									case "ulnone":
									case "ul0":
										result += "</u>";
										break;
									case "b":
										result += "<strong>";
										break;
									case "b0":
										result += "</strong>";
										break;
									case "i":
										result += "<em>";
										break;
									case "i0":
										result += "</em>";
										break;
									case "strike":
										result += "<s>";
										break;
									case "strike0":
										result += "</s>";
										break;
									case "colortbl":
										rdcolor = true;
										break;
									case "line":
										result += "\n";
										break;
									case "emdash":
										result += "&mdash;";
										break;
									case "endash":
										result += "&ndash;";
										break;
									case "lquote":
										result += "&lsquo;";
										break;
									case "rquote":
										result += "&rsquo;";
										break;
									case "ldblquote":
										result += "&ldquo;";
										break;
									case "rdblquote":
										result += "&rdquo;";
										break;
								}
							}
							if( rtf[cur].Equals( '\\' ) ) // start another tag
							{
								tag = true;
								tg = string.Empty;
							}
						}
						else if( rtf[cur].Equals( '{' ) )
						{
							brace++;
						}
						else if( rtf[cur].Equals( '}' ) )
						{
							brace--;
							if( brace < ftlevel )
							{
								rdfonts = false;
							}
						}
						else if( rtf[cur].Equals( '\\' ) )
						{
							tg = string.Empty;
							tag = true;
						}
						else //only interested in text
						{
							if( !tag && !char.IsControl( rtf[cur] ) )
							{
								result += rtf[cur];
							}
							else
							{
								switch( rtf[cur] )
								{
									case '<':
										tg += "&lt;";
										break;
									case '>':
										tg += "&gt;";
										break;
									case '&':
										tg += "&amp;";
										break;
									case '\"':
										tg += "&quot;";
										break;
									default:
										tg += rtf[cur];
										break;
								}
							}
						}
					}
					cur++;
				}

				result = result.TrimEnd( "\r\n".ToCharArray() );
				while( TagCounts.Font )
				{
					result += "</font>";
					TagCounts.Font = false;
				}
				while( TagCounts.Span )
				{
					result += "</span>";
					TagCounts.Span = false;
				}
				Regex re = new Regex( @"\<font[^>]+\>[ \n\r]*\</font\>", RegexOptions.Compiled );
				result = re.Replace( result, string.Empty );

				#region Write finished output to desktop for debugging
				//StreamWriter sw2 = File.CreateText( Environment.GetFolderPath( Environment.SpecialFolder.Desktop ) + "\\debugdone.txt" );
				//sw2.Write( result );
				//sw2.Close();
				#endregion

				if( result.Length > 0 )
				{
					if (bgcolor == "#ffffff")
						bgcolor="transparent";
					result = "<p style=\"background-color:" + bgcolor + "\">" + result + "</p>";
					return result;
				}
			}
			catch( Exception ex )
			{
				MessageBox.Show( ex.Message + "\n" + ex.StackTrace, @"Exception in " + ex.Source );
			}
			throw new Exception( "Error converting post to HTML." );
		}
	}


	internal class TagCounts
	{
		private static uint _font;
		private static uint _span;

		// static class
		private TagCounts(){}

		public static bool Font
		{
			get
			{
				return (_font > 0);
			}
			set
			{
				if( value )
					_font++;
				else
					_font--;
			}
		}
		public static bool Span
		{
			get
			{
				return (_span > 0);
			}
			set
			{
				if( value )
					_span++;
				else
					_span--;
			}
		}
	}
}

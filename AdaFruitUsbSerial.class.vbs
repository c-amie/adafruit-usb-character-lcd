	' AdafruitUsbSerial Application Programming Interface v1.0.3
	' © Chris Tilley | AmDi 1996 - 2014

	' DEPENDENCIES
	' ############

Class AdafruitUsbSerial

	private m_ForReading
	private m_SCREEN_OFF
	private m_SCREEN_ON
	private m_AUTO_SCROLL_ON
	private m_AUTO_SCROLL_OFF
	private m_CLEAR_SCREEN
	private m_SET_STARTUP_SPLASH
	private m_SET_CURSOR_POSITION
	private m_SET_CURSOR_HOME
	private m_SET_CURSOR_BACK
	private m_SET_CURSOR_FORWARD
	private m_SET_UNDERLINE_ON
	private m_SET_UNDERLINE_OFF
	private m_SET_BLINK_ON
	private m_SET_BLINK_OFF
	private m_SET_RGB

	private m_iPortNumber
	private m_bolDebug
	private m_bolAutoScroll
	private m_bolUnderlineCursor
	private m_bolBlinkCursor

	private m_fso

	private sub Class_Initialize
		m_ForReading			= 1
		m_SCREEN_OFF			= chr(254) & chr(70)
		m_SCREEN_ON				= chr(254) & chr(66)
		m_AUTO_SCROLL_ON		= chr(254) & chr(81)
		m_AUTO_SCROLL_OFF		= chr(254) & chr(82)
		m_CLEAR_SCREEN			= chr(254) & chr(88)
		m_SET_STARTUP_SPLASH	= chr(254) & chr(64)
		m_SET_CURSOR_POSITION	= chr(254) & chr(71)
		m_SET_CURSOR_HOME		= chr(254) & chr(72)
		m_SET_CURSOR_BACK		= chr(254) & chr(76)
		m_SET_CURSOR_FORWARD	= chr(254) & chr(77)
		m_SET_UNDERLINE_ON		= chr(254) & chr(74)
		m_SET_UNDERLINE_OFF		= chr(254) & chr(75)
		m_SET_BLINK_ON			= chr(254) & chr(83)
		m_SET_BLINK_OFF			= chr(254) & chr(84)
		m_SET_RGB				= chr(254) & chr(208)

		m_iPortNumber			= 1
		m_bolDebug				= false
		m_bolAutoScroll			= true
		m_bolUnderlineCursor	= false
		m_bolBlinkCursor		= false

		set m_fso 				= CreateObject("Scripting.FileSystemObject") 
	end sub

	private sub Class_Terminate
		set m_fso = nothing
	end sub

	' PROPERTIES
	public property get PortNumber
		PortNumber = m_iPortNumber
	end property

	public property let PortNumber(ByRef iIn)
		m_iPortNumber = iIn
	end property

	public property get Debug()
		Debug = m_bolDebug
	end property

	public property let Debug(ByRef bolIn)
		m_bolDebug = bolIn
	end property

	public property get AutoScroll()
		AutoScroll = m_bolAutoScroll
	end property

	public property let AutoScroll(ByRef bolIn)
		if (bolIn) then
			me.write(m_AUTO_SCROLL_ON)
		else
			me.write(m_AUTO_SCROLL_OFF)
		end if
		m_bolAutoScroll = bolIn
	end property

	public property get Underline()
		Underline = m_bolUnderlineCursor
	end property

	public property let Underline(ByRef bolIn)
		if (bolIn) then
			me.write(m_SET_UNDERLINE_ON)
		else
			me.write(m_SET_UNDERLINE_OFF)
		end if
		m_bolUnderlineCursor = bolIn
	end property

	public property get Blink()
		Blink = m_bolBlinkCursor
	end property

	public property let Blink(ByRef bolIn)
		if (bolIn) then
			me.write(m_SET_BLINK_ON)
		else
			me.write(m_SET_BLINK_OFF)
		end if
		m_bolBlinkCursor = bolIn
	end property

	' METHODS
	public sub clearScreen()
		me.write(m_CLEAR_SCREEN)
	end sub

	public sub screenOn()
		me.write(m_SCREEN_ON)
	end sub

	public sub screenOff()
		me.write(m_SCREEN_OFF)
	end sub

	public sub changeSplashScreen(ByVal strIn)
		strIn = Trim(strIn)
		strIn = Left(strIn, 32)
		me.write(m_SET_STARTUP_SPLASH)
		me.write(strIn)
	end sub

	public sub backlight(ByRef byteR, ByRef byteG, ByRef byteB)
		me.write(m_SET_RGB)
		me.write(chr(byteR))
		me.write(chr(byteG))
		me.write(chr(byteB))
	end sub

	public sub setCursorPosition(ByRef iX, ByRef iY)
		me.write(m_SET_CURSOR_POSITION)
		me.write(chr(iX))
		me.write(chr(iY))
	end sub

	public sub home()
		me.write(m_SET_CURSOR_HOME)
	end sub

	public sub back()
		me.write(m_SET_CURSOR_BACK)
	end sub

	public sub goBack(ByRef iIn)
		Dim i
		for i = 1 to iIn
			me.write(m_SET_CURSOR_BACK)
		next
	end sub

	public sub forward()
		me.write(m_SET_CURSOR_FORWARD)
	end sub

	public sub write(ByRef strIn)
		Dim serialWriter
		if (me.Debug) then
			wscript.echo strIn
		end if
		set serialWriter = m_fso.CreateTextFile("COM" & m_iPortNumber & ":",True)
			serialWriter.Write(strIn)
			serialWriter.Close()
		set serialWriter = nothing
	end sub

	public sub teletype(ByRef strIn, ByRef iDelayMs)
		Dim i
		Dim iLen
		iLen = Len(strIn)
		for i = 1 to iLen
			me.write(Mid(strIn, i, 1))
			WScript.Sleep(iDelayMs)
		next
	end sub

	public function testComPort(ByRef byteNumber)
		Dim serialWriter
		if (me.Debug) then
			wscript.echo "Attempting communications with COM" & byteNumber
		end if
		On Error Resume Next
			set serialWriter = m_fso.CreateTextFile("COM" & byteNumber & ":",True)
				serialWriter.Write("Initialising...")
				serialWriter.Close()
			set serialWriter = nothing
			if (err.number = 0) then
				testComPort = true
			else
				testComPort = false
			end if
		On Error Goto 0
	end function

End Class
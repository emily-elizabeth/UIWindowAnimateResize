#tag Module
Protected Module UIWindowAnimateResize
	#tag Method, Flags = &h0
		Sub UIWindowAnimateResize(win As DesktopWindow, width As Integer, height As Integer)
		  mAnimateResizeWidth = width
		  mAnimateResizeHeight = height
		  
		  Timer.CallLater 20, AddressOf UIWindowAnimateResizeTick, win
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UIWindowAnimateResizeCancel()
		  Timer.CancelCallLater AddressOf UIWindowAnimateResizeTick
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub UIWindowAnimateResizeTick(aWin As Variant)
		  // converted from https://stackoverflow.com/questions/1769317/animate-window-resize-width-and-height-c-sharp-wpf
		  
		  
		  DIM win As DesktopWindow = DesktopWindow(aWin)
		  
		  if (mAnimateResizeStop = 0) then
		    mAnimateResizeRatioHeight = ((win.Height - mAnimateResizeHeight) / 12) * -1
		    mAnimateResizeRatioWidth = ((win.Width - mAnimateResizeWidth) / 12) * -1
		  end if
		  mAnimateResizeStop = mAnimateResizeStop + 1
		  
		  win.Height = win.Height + mAnimateResizeRatioHeight
		  win.Width = win.Width + mAnimateResizeRatioWidth
		  
		  Timer.CallLater 20, AddressOf UIWindowAnimateResizeTick, win
		  
		  if (mAnimateResizeStop = 12) then
		    Timer.CancelCallLater AddressOf UIWindowAnimateResizeTick
		    win.Height = mAnimateResizeHeight
		    win.Width = mAnimateResizeWidth
		    
		    mAnimateResizeStop = 0
		    mAnimateResizeRatioHeight = 0
		    mAnimateResizeRatioWidth = 0
		    mAnimateResizeHeight = 0
		    mAnimateResizeWidth = 0
		  end if
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mAnimateResizeHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAnimateResizeRatioHeight As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAnimateResizeRatioWidth As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAnimateResizeStop As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAnimateResizeWidth As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule

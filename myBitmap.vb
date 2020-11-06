Class myBitmap

#Region " GrayScale "

    Public Shared Function ToGrayScale(ByVal srcImage As ImageSource) As ImageSource
        Dim grayBitmap As New FormatConvertedBitmap()
        grayBitmap.BeginInit()
        grayBitmap.Source = CType(srcImage, BitmapSource)
        grayBitmap.DestinationFormat = PixelFormats.Gray8
        grayBitmap.EndInit()
        Return grayBitmap
    End Function

#End Region

#Region " Bitmap conversion "

    Public NotInheritable Class BitmapConversion
        Private Sub New()
        End Sub

        '<System.Runtime.CompilerServices.Extension()> _
        Public Shared Function ToDrawingBitmap(bitmapsource As BitmapSource) As System.Drawing.Bitmap
            Using stream As New System.IO.MemoryStream()
                Dim enc As BitmapEncoder = New BmpBitmapEncoder()
                enc.Frames.Add(BitmapFrame.Create(bitmapsource))
                enc.Save(stream)

                Using tempBitmap = New System.Drawing.Bitmap(stream)
                    ' According to MSDN, one "must keep the stream open for the lifetime of the Bitmap."
                    ' So we return a copy of the new bitmap, allowing us to dispose both the bitmap and the stream.
                    Return New System.Drawing.Bitmap(tempBitmap)
                End Using
            End Using
        End Function

        '<System.Runtime.CompilerServices.Extension()> _
        Public Shared Function ToBitmapSource(bitmap As System.Drawing.Bitmap) As BitmapSource
            Using stream As New System.IO.MemoryStream()
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp)

                stream.Position = 0
                Dim result As New BitmapImage()
                result.BeginInit()
                ' According to MSDN, "The default OnDemand cache option retains access to the stream until the image is needed."
                ' Force the bitmap to load right now so we can dispose the stream.
                result.CacheOption = BitmapCacheOption.OnLoad
                result.StreamSource = stream
                result.EndInit()
                result.Freeze()
                Return result
            End Using
        End Function
    End Class

#End Region

#Region " Cursor "
    Public Shared Function ToCursor(ByVal source As DrawingImage, dimension As Size) As Cursor
        'BitmapSource
        Dim drawingVisual As DrawingVisual = New DrawingVisual()
        Dim drawingContext As DrawingContext = drawingVisual.RenderOpen()
        drawingContext.DrawImage(source, New Rect(New Point(0, 0), dimension))
        drawingContext.Close()
        Dim bmp As RenderTargetBitmap = New RenderTargetBitmap(CInt(dimension.Width), CInt(dimension.Height), 96, 96, PixelFormats.Pbgra32)
        bmp.Render(drawingVisual)

        'BitmapImage
        Dim encoder As PngBitmapEncoder = New PngBitmapEncoder()
        Dim memoryStream As IO.MemoryStream = New IO.MemoryStream()
        Dim bImg As BitmapImage = New BitmapImage()
        encoder.Frames.Add(BitmapFrame.Create(bmp))
        encoder.Save(memoryStream)
        memoryStream.Position = 0
        bImg.BeginInit()
        bImg.StreamSource = memoryStream
        bImg.EndInit()
        Return ToCursor(bImg.StreamSource)
        memoryStream.Close()
    End Function

    Public Shared Function ToCursor(ByVal Ico As System.Drawing.Icon) As Cursor
        Dim handle As New SafeIconHandle(Ico.Handle)
        Return System.Windows.Interop.CursorInteropHelper.Create(handle)
    End Function

    Public Shared Function ToCursor(ByVal myURI As Uri) As Cursor
        Dim imgStream As IO.Stream = Application.GetResourceStream(myURI).Stream
        If myURI.ToString.EndsWith("cur") Or myURI.ToString.EndsWith("ico") Then
            Return New Cursor(imgStream)
        Else
            Return ToCursor(imgStream)
        End If
    End Function

    Public Shared Function ToCursor(ByVal IOStream As IO.Stream) As Cursor
        Dim bit As New System.Drawing.Bitmap(IOStream)
        If bit.Size.Width > 64 Then bit = ResizeBitmap(bit, 64, 64)
        Dim curPtr As IntPtr = bit.GetHicon()
        Dim handle As New SafeIconHandle(curPtr)
        Return System.Windows.Interop.CursorInteropHelper.Create(handle)
    End Function

    Public Shared Function ToCursor(ByVal imgBitmap As BitmapImage) As Cursor
        Return ToCursor(imgBitmap.StreamSource)
    End Function

    Private Shared Function ResizeBitmap(bit As System.Drawing.Bitmap, nWidth As Integer, nHeight As Integer) As System.Drawing.Bitmap
        Dim result As New System.Drawing.Bitmap(nWidth, nHeight)
        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(DirectCast(result, System.Drawing.Image))
            g.DrawImage(bit, 0, 0, nWidth, nHeight)
        End Using
        Return result
    End Function

    Class SafeIconHandle
        Inherits Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function DestroyIcon(<[In]()> hIcon As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        Private Sub New()
            MyBase.New(True)
        End Sub

        Public Sub New(hIcon As IntPtr)
            MyBase.New(True)
            Me.SetHandle(hIcon)
        End Sub

        Protected Overrides Function ReleaseHandle() As Boolean
            Return DestroyIcon(Me.handle)
        End Function
    End Class

#End Region

#Region " Icon "

    Public Shared Function IconToImageSource(ByVal Ico As System.Drawing.Icon) As ImageSource
        Return System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(Ico.Handle, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
    End Function

    Public Shared Function UriToIcon(ByVal myURI As Uri) As System.Drawing.Icon
        Dim iconStream As System.IO.Stream = Application.GetResourceStream(myURI).Stream
        Return New System.Drawing.Icon(iconStream)
    End Function

    Public Shared Function UriToImageSource(ByVal myURI As Uri) As ImageSource
        Dim iconStream As System.IO.Stream = Application.GetResourceStream(myURI).Stream
        Return IconToImageSource(New System.Drawing.Icon(iconStream))
    End Function

#End Region

#Region " Merge Images "

    'Size 1=same as Image1, 2=half Image1; Place 1=left up, 2=right up
    Public Shared Function Merge(ByVal Image1 As ImageSource, ByVal Image2 As Uri, ByVal iSize As Integer, ByVal iPlace As Integer) As RenderTargetBitmap
        Dim frame2 As BitmapFrame = BitmapDecoder.Create(Image2, BitmapCreateOptions.None, BitmapCacheOption.OnLoad).Frames.First()
        Return Merge(Image1, frame2, iSize, iPlace)
    End Function

    Public Shared Function Merge(ByVal Image1 As ImageSource, ByVal Image2 As ImageSource, ByVal iSize As Integer, ByVal iPlace As Integer) As RenderTargetBitmap
        ' Gets the size of the images (I assume each image has the same size)
        Dim imageWidth As Integer = CInt(Image1.Width)
        Dim imageHeight As Integer = CInt(Image1.Height)

        Dim iLeft, iTop As Double
        Select Case iPlace
            Case 1
                iLeft = 0 : iTop = 0
            Case 2
                iLeft = imageWidth / iSize * 2 : iTop = 0
            Case 3
                iLeft = 0 : iTop = imageHeight / iSize * 2
            Case 4
                iLeft = imageWidth / iSize * 2 : iTop = imageHeight / iSize * 2
        End Select

        ' Draws the images into a DrawingVisual component
        Dim drawingVisual As New DrawingVisual()
        Using drawingContext As DrawingContext = drawingVisual.RenderOpen()
            drawingContext.DrawImage(Image1, New Rect(0, 0, imageWidth, imageHeight))
            drawingContext.DrawImage(Image2, New Rect(iLeft, iTop, imageWidth / iSize, imageHeight / iSize))
        End Using

        ' Converts the Visual (DrawingVisual) into a BitmapSource
        Dim bmp As New RenderTargetBitmap(imageWidth, imageHeight, 96, 96, PixelFormats.Pbgra32)
        bmp.Render(drawingVisual)
        Return bmp

        ' Creates a PngBitmapEncoder and adds the BitmapSource to the frames of the encoder
        'Dim encoder As New PngBitmapEncoder()
        'encoder.Frames.Add(BitmapFrame.Create(bmp))

        ' Saves the image into a file using the encoder
        'Using stream As IO.Stream = IO.File.Create(pathTileImage)
        ' encoder.Save(stream)
        'End Using
    End Function

#End Region

End Class
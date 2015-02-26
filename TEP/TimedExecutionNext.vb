Namespace TEP
    ''' <summary>
    ''' Execute a function each cycle until it returns TRUE then goes to the next function and again.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TimedExecutionNext
        Implements IDisposable

        Private mList_Function As Queue(Of NextFunction)
        Public Delegate Function NextFunction(ByVal cTime As Long, ByRef Tag() As Object) As Boolean
        Public Delegate Sub OnceFunction(ByRef Tag() As Object)
        Private mCurrentFunction As NextFunction
        Private TimeWatch As System.Diagnostics.Stopwatch
        Private mTotalTime As Long
        Private Started As Boolean
        Private Reset As Boolean

        '  Private mTag(0) As Object
        Private mTagList As Queue(Of Object())
        Private mCurrentTag() As Object

        Private mBeginFunction As Queue(Of OnceFunction)
        Private mEndFunction As Queue(Of OnceFunction)

        Public Sub New()
            mList_Function = New Queue(Of NextFunction)
            mTagList = New Queue(Of Object())

            mBeginFunction = New Queue(Of OnceFunction)
            mEndFunction = New Queue(Of OnceFunction)
        End Sub

        Public Sub Add(Nfunction As NextFunction, Tag() As Object)
            Add(Nfunction, Tag, AddressOf Empty_Start, AddressOf Empty_End)
        End Sub

        Public Sub Add(Nfunction As NextFunction, Tag() As Object, Begin As OnceFunction, [End] As OnceFunction)
            mList_Function.Enqueue(Nfunction)
            mTagList.Enqueue(Tag)
            mBeginFunction.Enqueue(Begin)
            mEndFunction.Enqueue([End])
        End Sub
        Private mOnceFunction As OnceFunction
        Public Sub Process()
            If Started = False Then
                If mList_Function.Count = 0 Then Exit Sub
                TimeWatch = System.Diagnostics.Stopwatch.StartNew()
                Reset = True
                Started = True
            End If

            If Reset = True Then
                If mList_Function.Count = 0 Then
                    Started = False
                    Exit Sub
                End If

                mCurrentFunction = mList_Function.Dequeue
                mCurrentTag = mTagList.Dequeue
                mOnceFunction = mBeginFunction.Dequeue()
                mOnceFunction(mCurrentTag)
                mTotalTime = 0
                TimeWatch.Restart()
                Reset = False
            End If

            mTotalTime = TimeWatch.ElapsedMilliseconds

            If mCurrentFunction(mTotalTime, mCurrentTag) = True Then
                mOnceFunction = mEndFunction.Dequeue()
                mOnceFunction(mCurrentTag)
                Reset = True
            End If
        End Sub

        Public Sub Reset_Time()
            mTotalTime = 0
            If Not TimeWatch Is Nothing Then TimeWatch.Restart()
        End Sub

        Public Sub Empty_Start(ByRef Tag() As Object)
        End Sub

        Public Sub Empty_End(ByRef Tag() As Object)
        End Sub
       
        Public ReadOnly Property Count As Integer
            Get
                Return mList_Function.Count
            End Get
        End Property
#Region "IDisposable Support"
        Private disposedValue As Boolean

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    mList_Function.Clear()
                    mList_Function = Nothing
                    If Not TimeWatch Is Nothing Then
                        TimeWatch.Stop()
                        TimeWatch = Nothing
                    End If

                    mTagList.Clear()
                    mTagList = Nothing
                End If

            End If
            Me.disposedValue = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class
End Namespace

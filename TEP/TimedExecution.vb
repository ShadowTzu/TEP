Namespace TEP
    ''' <summary>
    '''  Execute a function each cycle until the time runs out and move to the next function and again.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TimedExecution
        Implements IDisposable

        Private Structure struct_TempoConfig
            Public Start_Time As Integer
            Public Duration As Integer
        End Structure

        Private mList_Function As Queue(Of TempoFunction)
        Private mTempoConfig As Queue(Of struct_TempoConfig)

        Public Delegate Sub TempoFunction(ByVal cTime As Long, mTime As Long)

        Private mCurrentFunction As TempoFunction
        Private mCurrentConfig As struct_TempoConfig
        Private mTotalTime As Long

        Private TimeWatch As System.Diagnostics.Stopwatch
        Private Started As Boolean
        Private Reset As Boolean

        Public Sub New()
            mList_Function = New Queue(Of TempoFunction)
            mTempoConfig = New Queue(Of struct_TempoConfig)
        End Sub

        Public Sub Add(Tfunction As TempoFunction, Start_time As Integer, Duration As Integer)
            mList_Function.Enqueue(Tfunction)
            Dim Config As struct_TempoConfig
            Config.Start_Time = Start_time
            Config.Duration = Duration
            mTempoConfig.Enqueue(Config)
        End Sub

        Public Sub Clear()
            Reset = False
            Started = False
            mList_Function.Clear()
            mTempoConfig.Clear()
        End Sub

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
                mCurrentConfig = mTempoConfig.Dequeue
                mTotalTime = 0
                TimeWatch.Restart()
                Reset = False
            End If

            mTotalTime = TimeWatch.ElapsedMilliseconds

            If mTotalTime >= mCurrentConfig.Start_Time Then
                mCurrentFunction(mTotalTime - mCurrentConfig.Start_Time, mCurrentConfig.Duration)
            End If

            If mTotalTime >= mCurrentConfig.Start_Time + mCurrentConfig.Duration Then
                Reset = True
            End If
        End Sub

#Region "Destructor"
        Private disposedValue As Boolean

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    mList_Function.Clear()
                    mList_Function = Nothing
                    mTempoConfig.Clear()
                    mTempoConfig = Nothing
                    If Not TimeWatch Is Nothing Then
                        TimeWatch.Stop()
                        TimeWatch = Nothing
                    End If
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

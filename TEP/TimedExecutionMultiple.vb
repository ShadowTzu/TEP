Namespace TEP
    ''' <summary>
    '''  Execute any function at the same time in each cycle until their time is up.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TimedExecutionMultiple
        Implements IDisposable

        Private Structure struct_TempoConfig
            Public Start_Time As Integer
            Public Duration As Integer
            Public TotalTime As Long
            Public Started As Boolean
            Public [Loop] As Boolean
            Public TimeWatch As System.Diagnostics.Stopwatch
            Public Tag() As Object
        End Structure

        Private mList_Function As List(Of TempoFunction)
        Private mTempoConfig As List(Of struct_TempoConfig)
        Private mRemoveList(100) As Integer
        Private mMaxRemove As Integer
        Public Delegate Sub TempoFunction(cTime As Long, mTime As Long, Tag() As Object)

        Private mCurrentFunction As TempoFunction
        Private mCurrentConfig As struct_TempoConfig

        Public Sub New()
            mList_Function = New List(Of TempoFunction)
            mTempoConfig = New List(Of struct_TempoConfig)

            For i As Integer = 0 To mRemoveList.Length - 1
                mRemoveList(i) = -1
            Next
        End Sub

        Public Sub Add(Tfunction As TempoFunction, Start_time As Integer, Duration As Integer, Tag() As Object, Optional [Loop] As Boolean = False)
            mList_Function.Add(Tfunction)
            Dim Config As struct_TempoConfig
            Config.Start_Time = Start_time
            Config.Duration = Duration
            Config.Started = False
            Config.TimeWatch = Nothing
            Config.Tag = Tag
            Config.Loop = [Loop]
            mTempoConfig.Add(Config)
        End Sub
        Private mI As Integer
        Public Sub Process()
            If mList_Function.Count = 0 Then Exit Sub
            For mI = 0 To mList_Function.Count - 1
                mCurrentConfig = mTempoConfig(mI)
                mCurrentFunction = mList_Function(mI)

                If mCurrentConfig.Started = False Then
                    mCurrentConfig.TimeWatch = System.Diagnostics.Stopwatch.StartNew()
                    mCurrentConfig.Started = True
                End If

                mCurrentConfig.TotalTime = mCurrentConfig.TimeWatch.ElapsedMilliseconds

                If mCurrentConfig.TotalTime >= mCurrentConfig.Start_Time AndAlso mCurrentConfig.TotalTime <= (mCurrentConfig.Start_Time + mCurrentConfig.Duration) Then
                    mCurrentFunction(mCurrentConfig.TotalTime - mCurrentConfig.Start_Time, mCurrentConfig.Duration, mCurrentConfig.Tag)


                ElseIf mCurrentConfig.TotalTime >= mCurrentConfig.Start_Time + mCurrentConfig.Duration Then
                    If mCurrentConfig.Loop Then
                        mCurrentConfig.TotalTime = 0
                        mCurrentConfig.Started = False
                    Else
                        'remove
                        mMaxRemove += 1
                        If mMaxRemove > mRemoveList.Length - 1 Then
                            ReDim Preserve mRemoveList(mMaxRemove)
                        End If
                        mRemoveList(mMaxRemove) = mI
                    End If

                End If
                mTempoConfig(mI) = mCurrentConfig

            Next

            For i As Integer = mMaxRemove - 1 To 0 Step -1
                If mRemoveList(i) > -1 Then
                    mList_Function.RemoveAt(mRemoveList(i))
                    mTempoConfig.RemoveAt(mRemoveList(i))
                    mRemoveList(i) = -1
                End If
            Next
            mMaxRemove = -1
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
                    Erase mRemoveList
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

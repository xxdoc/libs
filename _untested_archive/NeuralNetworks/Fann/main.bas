Option Explicit

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Sub GetMem8 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Double)

Private Sub Form_Load()
  If LoadLibraryW(StrPtr(App.Path & "\fanndouble.dll")) = 0 Then MsgBox "loading of fann-dll was not successful": Unload Me
End Sub

Private Sub Command1_Click()
Dim a(0 To 2) As Long
Dim ann As Long, train_data As Long
Dim inputs(0 To 1) As Double
Dim pDblArr As Long, DblVal As Double
Dim i As Integer, j As Integer
 
    a(0) = 2 'input neurons
    a(1) = 3 'hidden neurons
    a(2) = 1 'output neurons
    ann = fann_create_standard_array(UBound(a) + 1, a(0))
 
    train_data = fann_read_train_from_file(App.Path & "\xor.data")

    fann_set_activation_steepness_hidden ann, 0.5
    fann_set_activation_steepness_output ann, 0.5

    fann_set_activation_function_hidden ann, FANN_SIGMOID
    fann_set_activation_function_output ann, FANN_SIGMOID

    fann_set_train_stop_function ann, FANN_STOPFUNC_BIT
    fann_set_bit_fail_limit ann, 0.0001

    fann_init_weights ann, train_data
    fann_train_on_data ann, train_data, 500000, 1000, 0.0001

    For i = 0 To 1: For j = 0 To 1
        inputs(0) = i
        inputs(1) = j

        pDblArr = fann_run(ann, inputs(0))
        GetMem8 pDblArr, DblVal
        
        Text1.Text = Text1.Text & inputs(0) & " Xor " & inputs(1) & " = " & Round(DblVal, 9) & vbCrLf
    Next j, i
 
    fann_destroy ann
    Text1.Text = Text1.Text & vbCrLf
End Sub
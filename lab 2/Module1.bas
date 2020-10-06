Attribute VB_Name = "Module1"
Public Function steps_count(min_x, max_x, step)
    '������� ���������� ������� ����� ����� ����������� � �������
    'min_x, max_x - �������/������ ������� �������
    'step - ���
    
    '� ������ ������ ������ �������� ������� ����
    '���� ���� max_x < min_x ������� ����� �����
    '������������� �����
    steps_count = Fix((max_x - min_x) / step)

End Function

Public Function check_value_interval(a, b, value)
    '������� ������������ �������� ��������� ���������� �������� � ��������
    ' a,b - ������� ���������
    ' value - �������� ��� �������� �������������� ��������
    If value > a And value < b Then
        check_value_interval = True
    Else
        check_value_interval = False
    End If

End Function


Public Function plot_data_even_int(arr, a, b)
    '�������� ������������ ������������� � ��������� ����������
    'a,b - ��������� ��������� ��� ������������ �������������
    'arr - ������ ��� �������� ��������� ��������� ��������
    
    result_size = arr.Count
    Dim result()
    ReDim result(1 To result_size, 1 To 1)
    Dim cv As Integer
    
    For i = 1 To result_size

        If arr(i) < a Then
            result(i, 1) = 0
        ElseIf arr(i) >= b Then
            result(i, 1) = 1
        Else
            result(i, 1) = (arr(i) - a) / (b - a)
        End If
        
    Next i
    
    plot_data_even_int = result
End Function

Public Function even_dispersion(a, b)
    even_dispersion = ((b - a) ^ 2) / 12
End Function

Public Function poisson_median(lambda)
    poisson_median = lambda + (1 / 3) - (0.02 / lambda)
End Function

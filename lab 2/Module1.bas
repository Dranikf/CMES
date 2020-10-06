Attribute VB_Name = "Module1"
Public Function steps_count(min_x, max_x, step)
    'функция подсчитает сколько ЦЕЛЫХ шагов поместиться в отрезке
    'min_x, max_x - верхний/нижний границы отрезка
    'step - шаг
    
    'в данном случае просто отбросим дробную чать
    'даже если max_x < min_x получим чтото вроде
    'отрицательных шагов
    steps_count = Fix((max_x - min_x) / step)

End Function

Public Function check_value_interval(a, b, value)
    'функция осуществляет проверку попадания некоторого значения в интервал
    ' a,b - границы интервала
    ' value - значение для которого осуществляется проверка
    If value > a And value < b Then
        check_value_interval = True
    Else
        check_value_interval = False
    End If

End Function


Public Function plot_data_even_int(arr, a, b)
    'значения равномерного распределения с заданными праметрами
    'a,b - привычные параметры для равномерного распределения
    'arr - массив для которого требуется расчитать значения
    
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

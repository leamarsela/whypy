Attribute VB_Name = "Clay03"
Option Explicit

public diameter as Double 'unit in m
public modulus as double 'unit in m
public inertia as double 'unit in m4
public pt as double 'unit in kN
public i as integer
public j as integer
public k as integer
public l as integer
public temp as double
public sum as double
public id as integer
public mi as integer

dim wsInput as worksheet
dim wsReport as worksheet
dim wsCalc as worksheet

const pi = 3.14159265358979
const consJ = 0.5

function valDiameter()

    set wsInput = worksheets("Input")
    valDiameter = wsInput.cells(1,2)

end function

function valModulus()

    set wsInput = worksheets("Input")
    valModulus = wsInput.cells(2,2)

end function

function valInertia()

    set wsInput = worksheets("Input")
    valInertia = wsInput.cells(3,2)

end function

function valStiffness()

    valStiffness = valModulus() * valInertia()

end function

function valPt()

    set wsInput = worksheets("Input")
    valPt = wsInput.cells(1,6)

end function

function deltaPt()

    deltaPt = 5

end function

function valSegmen()

    set wsInput = worksheets("Input")
    valSegmen = wsInput.cells(2,6)

end function

function nData()

    set wsInput = worksheets("Input")
    nData = wsInput.application.range(cells(7,1), cells(7,1).end(xlDown)).rows.count

end function

function zoI(nI as integer)

    set wsInput = worksheets("Input")
    zoI = wsInput.application.range(cells(7,2), cells(7,2).end(xlDown)).cells(nI)

end function

function ziI(nI as integer)

    set wsInput = worksheets("Input")
    ziI = wsInput.application.range(cells(7,3), cells(7,3).end(xlDown)).cells(nI)

end function

function gammaI(nI as integer)

    set wsInput = worksheets("Input")
    gammaI = wsInput.application.range(cells(7,4), cells(7,4).end(xlDown)).cells(nI)

end function

function cuI(nI as integer)

    set wsInput = worksheets("Input")
    cuI = wsInput.application.range(cells(7,5), cells(7,5).end(xlDown)).cells(nI)

end function

function nodal()

    nodal = valSegmen() + 1

end function

function numSubB()

    numSubB = 2*nodal()

end function

function deltaX()

    deltaX = ziI(nData())/valSegmen()

end function

function numPoint()

    numPoint = 20

end function

sub interpolate(valXi, valYi, xVal, yVal)       'untuk interpolasi langrange         
                                            
    for k = 0 to (nodal()-1)
        for l = 0 to (numPoint()-1)
        sum = 0
            for i = 0 to (numPoint()-1)
                temp = yVal(k,i)
                for j = 0 to (numPoint()-1)
                    if(i<>j) then
                        temp = temp*(valXi(k,l) - xVal(k,j))/(xVal(k,i) - xVal(k,j))
                    end if
                next j
                sum = sum + temp
                valYi(k,l) = sum
            next i
        next l
    next k
    
end sub


sub softClay()

redim nDepth(nodal()-1)
redim gammas(nodal()-1)
redim cus(nodal()-1)
redim layer(nodal()-1)
redim gammaDepth(nodal()-1)
redim cuDepth(nodal()-1)
redim gammaAvg(nodal()-1)
redim cuAvg(nodal()-1)
redim pult1(nodal()-1)
redim pult2(nodal()-1)
redim pult(nodal()-1)
redim epsilon50(nodal()-1)
redim valY50(nodal()-1)
redim valY(nodal()-1, numpoint()-1)
redim valP(nodal()-1, numpoint()-1)
redim km(nodal()-1, numpoint()-1)
redim kmI(nodal()-1, numpoint()-1)
redim valPtI(deltaPt()-1, numpoint()-1)
redim valCI(deltaPt()-1, numpoint()-1)
redim valA(nodal()-1, numpoint()-1)
redim valB(numSubB()-1, numpoint()-1)
redim valD(3, nunpoint()-1)
redim valV(nodal()+4, numpoint()-1)


    nDepth(0) = 0                     'menghitung kedalaman tiap titik nodal
    temp = 0                      
    for i = 1 to (nodal()-1)
        nDepth(i) = temp + deltaX()
        temp = nDepth(i) 
    next i

    for i = 0 to (nodal()-1)          'menentukan nilai c dan gamma untuk setiap lapisan
        for j = 0 to (nData())
            if i = 0 then
                cus(i) = cuI(0)
                gammas(i) = gammaI(0)
            elseif i > 0 and i < (nodal()-1) then
                if nDepth(i) > zoI(j) and nDepth(i) < ziI(j) then
                    cus(i) = cuI(j)
                    gammas(i) = gammaI(j)
                end if
            elseif i = (nodal()-1) then
                cus(i) = cuI(nData())
                gammas(i) = gammaI(nData())
            end if
        next j
    next i

    
    temp = 0
    cuDepth(0) = 0                    'menghitung cu x tinggi layer                   
    for i = 1 to (nodal()-1)
        if i = 0 then
            layer(i) = nDepth(1)
        elseif i = (nodal()-1) then
            layer(i) = nDepth(nodal() - 1) - nDepth(nodal() - 2)
        else
            layer(i) = nDepth(i+1) - nDepth(i)
        end if

        cuDepth(i) = temp + (layer(i)*cus(i))
        temp = cuDepth(i)
    next i

    temp = 0
    gammaDepth(0) = 0                 'menghitung gamma x tinggi layer           
    for i = 0 to (nodal()-1)
        if i = 0 then
            layer(i) = nDepth(1)
        elseif i = (nodal()-1) then
            layer(i) = nDepth(nodal() - 1) - nDepth(nodal() - 2)
        else
            layer(i) = nDepth(i+1) - nDepth(i)
        end if

        gammaDepth(i) = temp + (layer(i)*gammas(i))
        temp = gammaDepth(i)
    next i

    for i = 0 to (nodal()-1)          'menghitung gamma average and cu average untuk setiap titik nodal
        if i = 0 then
            cuAvg(i) = cus(0)
            gammaAvg(i) = gammas(0)
        else
            cuAvg(i) = cuDepth(i) / nDepth(i)
            gammaAvg(i) = gammaDepth(i) / nDepth(i)
        end if
    next i

    for i = 0 to (nodal()-1)          'menghitung nilai pult untuk setiap titik nodal
        pult1(i) = 9*cuAvg(i)*valDiameter()
        pult2(i) = (3 + (gammaAvg(i)*nDepth(i)/cuAvg(i)) + (consJ*nDepth(i)/valDiameter() ))*cuAvg(i)*valDiameter()

        if (pult1(i) <= pult2(i)) then
            pult(i) = pult1(i)
        else
            pult(i) = pult2(i)
        end if
    next i

    for i = 0 to (nodal()-1)          'menentukan nilai epsilon50 untuk setiap titik nodal
        if cuAvg(i) < 48 then
            epsilon50(i) = 0.02
        elseif cuAvg(i) >= 48 and cuAvg(i) < 96 then
            epsilon50(i) = 0.01
        elseif cuAvg(i) => 96 and cuAvg(i) < 192 then
            epsilon50(i) = 0.005
        else
            epsilon50(i) = 0.005
        end if
    next i

    for i = 0 to (nodal()-1)          'menghitung nilai y50 untuk setiap titik nodal
        valY50(i) = 2.5 * valDiameter() * epsilon50(i)
    next i

    for i = 0 to (nodal()-1)          'menghitung nilai y untuk setiap titik nodal
        for j = 0 to (numPoint()-1)
            valY(i,j) = valY50(i) * j
        next j
    next i

    for i = 0 to (nodal()-1)          'menghitung nilai p untuk setiap titik nodal
        for j = 0 to (numPoint()-1)
            if j <= 8 then
                valP(i,j) = 0.5 * ((valY(i,j)/valY50(i))^(1/3))*pult(i)
            else
                valP(i,j) = pult(i)
            end if
        next j
    next i

    for i = 0 to (nodal()-1)          'menghitung nilai km untuk setiap titik nodal
        for j = 0 to (numPoint()-1)
            if j >= 8 then
                km(i,j) = 0
            else
                km(i,j) = (valP(i, j+1) - valP(i, j)) /(valY(i, j+1) - valY(i, j))
            end if
        next j
    next i

    for i = 0 to (deltaPt()-1)
        for j = 0 to (numpoint()-1)

            valPtI(i) = valPt() / (deltaPt() - i)                           'menghitung delta Pt
            valCI(i) = 2*valPtI(i)*(deltaX()^3)/(valStiffness())            'menghitung nilai C

            kmI(i,j) = km(0,0)
            do while abs(kmI(i,j) - km(i,j)) > 0.001        
                for i = 0 to (nodal()-1)                                    'menghitung nilai A
                    for j = 0 to (numPoint()-1)
                        km(i,j) = kmI(i,j)
                        id = (nodal()-1)-i
                        valA(id,j) = km(i,j) * (deltaX())^4 / (valStiffness())
                    next j
                next i

                for i = 0 to (numSubB()-1)                                  'menghitung nilai B
                    for j = 0 to (numpoint()-1)
                        if i = 0 then
                            valB(i,j) = 2/(valA(i,j)+2)
                        elseif i = 1 then
                            valB(i,j) = 2*valB(i-1,j)
                        elseif i = 2 then
                            valB(i,j) = 1/(5+(valA(i-1,j)-2*valB(i-1,j)))
                        elseif i > 2 then
                            if (i Mod 2 > 0) then
                                mi = (i-1)/2
                                valB(2*mi+1,j) = valB(2*mi,j)*(4-valB(2*mi-1,j))
                            elseif (i Mod 2 = 0) then
                                mi = (i/2)
                                valB(2*mi,j) = 1/(6+valA(mi,j)-valB(2*mi-4,j)-(valB(2*mi-1,j)*(4-valB(2*mi-3,j))))
                            end if
                        end if
                    next j
                next i

                for i = 1 to 3                                              'menghitung nilai D
                    for j = 0 to (numpoint()-1)
                        if (i = 1) then
                            valD(i,j) = 1/valB(2*(nodal()-1),j)
                        elseif (i = 2) then
                            valD(i,j) = valD(i-1,j)*valB(2*(nodal()-1)+1,j)-(valB(2*(nodal()-1)-2,j)*(2-valB(2*(nodal()-1)-3,j)))-2
                        elseif (i = 3) then
                            valD(i,j) = valD(i-2,j)-valB(2*(nodal()-1)-4,j)-valB(2*(nodal()-1)-1,j)*(2-valB(2*(nodal()-1)-3,j))
                        end if
                    next j
                next i

                





            loop


        next j
    next i
    

    

end sub

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
public n as integer
public m as integer

dim wsInput as worksheet
dim wsReport as worksheet
dim wsCalc as worksheet

dim deltaPt as integer
dim point as integer

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
    nData = wsInput.application.range(cells(6,1), cells(6,1).end(xlDown)).rows.count

end function

function zoI(nI as integer)

    set wsInput = worksheets("Input")
    zoI = wsInput.application.range(cells(6,2), cells(6,2).end(xlDown)).cells(nI)

end function

function ziI(nI as integer)

    set wsInput = worksheets("Input")
    ziI = wsInput.application.range(cells(6,3), cells(6,3).end(xlDown)).cells(nI)

end function

function gammaI(nI as integer)

    set wsInput = worksheets("Input")
    gammaI = wsInput.application.range(cells(6,4), cells(6,4).end(xlDown)).cells(nI)

end function

function cuI(nI as integer)

    set wsInput = worksheets("Input")
    cuI = wsInput.application.range(cells(6,5), cells(6,5).end(xlDown)).cells(nI)

end function

function nodal()

    nodal = valSegmen() + 1

end function

function numSubB()

    numSubB = (2*(nodal()-1))+1

end function

function deltaX()

    deltaX = ziI(nData() - 1)/valSegmen()

end function

function numPoint()

    numPoint = 20

end function

function interpolate(valYi as double, xVal(), yVal())

dim sum as double
dim L as double

    sum = 0
    for i = 1 to numPoint()
        L = 1
        for j = 1 to numPoint()
            if(i<>j) then
                L = L*(valYi - xVal(j,1))/(xVal(i,1) - xVal(j,1))
            end if
        next j
        sum = sum + L*yVal(i,1)
    next i

    interpolate = sum

end function


sub softClay()

redim nDepth(nodal())
redim nDepthI(nodal())
redim gammas(valSegmen())
redim cus(valSegmen())
redim layer(valSegmen())
redim gammaDepth(valSegmen())
redim gammaAvg(valSegmen())
redim pult1(valSegmen())
redim pult2(valSegmen())
redim pult(valSegmen())
redim epsilon50(valSegmen())
redim valY50(valSegmen())
redim valY(valSegmen(), numPoint())
redim valP(valSegmen(), numPoint())
redim km(valSegmen(), numPoint())
redim kmI(valSegmen(), numPoint())
redim valPtI(valSegmen())
redim valCI(valSegmen())
redim valA(valSegmen())
redim valB(valSegmen())
dim valMi as double


    nDepth(0) = 0
    nDepthI(0) = 0                      'menghitung kedalaman tiap titik nodal
    for i = 1 to (nodal())
        nDepth(i) = deltaX() + nDepthI(0)
        nDepthI(0) = nDepth(i) 
    next i

    for i = 0 to (valSegmen())      'menentukan nilai c dan gamma untuk setiap lapisan
        for j = 0 to (nData() - 1)
            if i = 0 then
                cus(i) = cuI(0)
                gammas(i) = gammaI(0)
            elseif i > 0 and i < (valSegmen()) then
                if nDepth(i) > zoI(j) and nDepth(i) < ziI(j) then
                    cus(i) = cuI(j)
                    gammas(i) = gammaI(j)
                end if
            elseif i = (valSegmen()) then
                cus(i) = cuI(nData() - 1)
                gammas(i) = gammaI(nData() - 1)
            end if
        next j
    next i

    gammaDepth(0) = 0                   'menghitung gamma x tinggi layer
    for i = 0 to (valSegmen())
        if i = 0 then
            layer(i) = nDepth(1)
        elseif i = (valSegmen()) then
            layer(i) = nDepth(valSegmen() - 1) - nDepth(valSegmen() - 2)
        else
            layer(i) = nDepth(i+1) - nDepth(i)
        end if

        gammaDepth(i) = gammaDepth(0) + (layer(i)*gammas(i))
        gammaDepth(0) = gammaDepth(i)
    next i

    for i = 0 to (valSegmen())          'menghitung gamma average
        if i = 0 then
            gammaAvg(i) = gammas(0)
        else
            gammaAvg(i) = gammaDepth(i) / nDepth(i)
        end if
    next i

    for i = 0 to (valSegmen())          'menghitung nilai pult
        pult1(i) = 9*cus(i)*valDiameter()
        pult2(i) = (3 + (gammaAvg(i)*nDepth(i)/cus(i)) + (consJ*nDepth(i)/valDiameter() ))*cus(i)*valDiameter()

        if (pult1(i) <= pult2(i)) then
            pult(i) = pult1(i)
        else
            pult(i) = pult2(i)
        end if
    next i

    for i = 0 to (valSegmen())          'menentukan nilai epsilon50
        if cus(i) < 48 then
            epsilon50(i) = 0.02
        elseif cus(i) >= 48 and cus(i) < 96 then
            epsilon50(i) = 0.01
        elseif cus(i) => 96 and cus(i) < 192 then
            epsilon50(i) = 0.005
        else
            epsilon50(i) = 0.005
        end if
    next i

    for i = 0 to (valSegmen())          'menghitung nilai y50
        valY50(i) = 2.5 * valDiameter() * epsilon50(i)
    next i

    for i = 0 to (valSegmen())          'menghitung nilai y
        for j = 0 to numPoint()
            valY(i,j) = valY50(i) * j
        next j
    next i

    for i = 0 to (valSegmen())          'menghitung nilai p
        for j = 0 to numPoint()
            if j <= 8 then
                valP(i,j) = 0.5 * ((valY(i,j)/valY50(i))^(1/3))*pult(i)
            else
                valP(i,j) = pult(i)
            end if
        next j
    next i

    for i = 0 to (valSegmen())          'menghitung nilai km
        for j = 0 to (numPoint())
            if j >= 8 then
                km(i,j) = 0
            else
                km(i,j) = (valP(i, j+1) - valP(i, j)) /(valY(i, j) - valY(i, j))
            end if
        next j
    next i

    for i = 0 to deltaPt()              'menghitung delta Pt dan nilai C

        valPtI(i) = valPt() / (deltaPt() - i)
        valCI(i) = 2*valPtI(i)*(deltaX()^3)/(valStiffness())

    next i
    

    for i = 0 to (numSubB())
        for j = 0 to (numPoint())
            do while abs(km(i,j) - kmI(i,j)) > 0.001
                kmI(i,j) = km(i,j)
                valA(i) = kmI(i,j) * (deltaX())^4 / (valStiffness())
                if i = 0 then
                    valB(i) = 2/(valA(valSegmen())+2)
                elseif i = 1 then
                    valB(i) = 2*valB(i-1)
                elseif i = 2 then
                    valB(i) = 1/(5+valA(valSegmen()-(i-1)) - (2*valB(i-1)))
                elseif i>2 then
                    if i Mod 2>0 then
                        valMi = (i-1)/2
                        valB(i) = valB(2*valMi)*(4-valB((2*valMi)-1))
                    elseif m Mod 2 = 0 then
                        valMi = (i/2)
                        valB(i) = 1/(6+(valA(valSegmen()-valMi)-(valB(2*valMi-4))-(valB(2*valMi-1)*(4-valB(2*valMi-3))))
                    end if
                end if
            loop
        next j
    next i


    

end sub

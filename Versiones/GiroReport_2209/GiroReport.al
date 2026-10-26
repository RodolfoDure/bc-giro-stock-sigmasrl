report 50105 "Giro Stock Report"
{
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;
    Caption = 'Giro Stock Report';
    DefaultLayout = Excel;
    ExcelLayout = './src/Reports/Layouts/GiroStockReport.xlsx';

    dataset
    {
        dataitem(Item; Item)
        {
            RequestFilterFields = "No.", "Item Category Code"; //, "Global Dimension 1 Code"

            // Columnas del producto
            column(Familia; "Item Category Code") { }
            column(Nombre_Familia; vNombreFamilia) { }
            column(Codigo; "No.") { }
            column(Descripcion; Description) { }
            // Totales e Índices principales
            column(Stock; vStockALaFecha) { }
            column(Ventas; vVentasPeriodo) { }
            column(IND; vIND) { }
            // Columnas de antigüedad según tu script
            column(Dias_30; vN_D30) { }
            column(Dias_60; vN_D60) { }
            column(Dias_90; vN_D90) { }
            column(Dias_180; vN_D180) { }
            column(Dias_360; vN_D360) { }
            column(Mas_1Anho; vN_1ANO_Mas) { }
            column(Anhos_5; vN_2ANO_Mas) { }
            column(Mas_5Anhos; vN_3ANO_Mas) { }
            // Columnas finales solicitadas
            column(Sin_Ventas_12Meses; vSinVenta12M) { }
            column(Sin_Ventas_36Meses; vSinVenta36M) { }
            column(Total_Stock_Costo; vTotalStockCosto) { }

            trigger OnAfterGetRecord()
            var
                ItemLedgerEntry: Record "Item Ledger Entry";
            begin
                ClearCalculations();

                if ItemCategory.Get("Item Category Code") then
                    vNombreFamilia := ItemCategory.Description;

                // 1. Calcular Stock a la fecha final del filtro
                ItemLedgerEntry.Reset();
                ItemLedgerEntry.SetCurrentKey("Item No.", "Posting Date");
                ItemLedgerEntry.SetRange("Item No.", Item."No.");
                ItemLedgerEntry.SetRange("Posting Date", 0D, FechaFinalFiltro);
                if ItemLedgerEntry.FindSet() then begin
                    repeat
                        vStockALaFecha += ItemLedgerEntry.Quantity;
                    until ItemLedgerEntry.Next() = 0;
                end;

                // Cálculo seguro del costo monetario
                vCostoALaFecha := vStockALaFecha * Item."Unit Cost";

                // 2. Calcular Ventas dentro del rango solicitado
                ItemLedgerEntry.Reset();
                ItemLedgerEntry.SetCurrentKey("Item No.", "Posting Date");
                ItemLedgerEntry.SetRange("Item No.", Item."No.");
                ItemLedgerEntry.SetRange("Posting Date", FechaInicialFiltro, FechaFinalFiltro);
                ItemLedgerEntry.SetRange("Entry Type", "Item Ledger Entry Type"::Sale);
                /* ItemLedgerEntry.SetFilter("Document Type", '%1|%2|%3|%4',
                     "Item Ledger Document Type"::"Sales Invoice",
                     "Item Ledger Document Type"::"Sales Shipment",
                     "Item Ledger Document Type"::"Sales Return Receipt",
                     "Item Ledger Document Type"::"Sales Credit Memo");*/
                if ItemLedgerEntry.FindSet() then begin
                    repeat
                        vVentasPeriodo += (ItemLedgerEntry.Quantity * -1);
                    until ItemLedgerEntry.Next() = 0;
                end;

                // 3. Calcular Balde de compras últimas 3 meses
                if FechaFinalFiltro <> 0D then begin
                    FechaInicio3M := CalcDate('<-3M>', FechaFinalFiltro);

                    ItemLedgerEntry.Reset();
                    ItemLedgerEntry.SetCurrentKey("Item No.", "Posting Date");
                    ItemLedgerEntry.SetRange("Item No.", Item."No.");
                    ItemLedgerEntry.SetRange("Posting Date", FechaInicio3M, FechaFinalFiltro);
                    ItemLedgerEntry.SetRange("Entry Type", "Item Ledger Entry Type"::Purchase);
                    if ItemLedgerEntry.FindSet() then begin
                        repeat
                            vCompras3M += ItemLedgerEntry.Quantity;
                        until ItemLedgerEntry.Next() = 0;
                    end;
                end;

                // 4. Cálculo del IND
                if (vStockALaFecha - vCompras3M) <= 0 then begin
                    vIND := 12;
                end else begin
                    vIND := vVentasPeriodo / (vStockALaFecha - vCompras3M);
                end;

                // 5. Condición prioritaria de inactividad basada en el rango elegido por el usuario
                if (vCompras3M = 0) and (vVentasPeriodo = 0) then begin
                    vTotalStockCosto := vCostoALaFecha;

                    if (FechaInicialFiltro <> 0D) and (FechaFinalFiltro <> 0D) then begin
                        // RUTA 1: Si el rango seleccionado en pantalla es de 1 año o menos (aprox. 366 días)
                        if (FechaFinalFiltro - FechaInicialFiltro) <= 366 then begin
                            vSinVenta12M := vCostoALaFecha;
                            vSinVenta36M := 0;
                        end else begin
                            // RUTA 2: Si el rango en pantalla es ampliado (Análisis histórico)
                            // 1. Evaluamos primero la inactividad extrema (36 Meses hacia atrás)
                            FechaInicio36M := CalcDate('<-36M>', FechaFinalFiltro);
                            ItemLedgerEntry.Reset();
                            ItemLedgerEntry.SetCurrentKey("Item No.", "Posting Date");
                            ItemLedgerEntry.SetRange("Item No.", Item."No.");
                            ItemLedgerEntry.SetRange("Posting Date", FechaInicio36M, FechaFinalFiltro);
                            ItemLedgerEntry.SetRange("Entry Type", "Item Ledger Entry Type"::Sale);
                            ItemLedgerEntry.SetFilter("Document Type", '%1|%2|%3|%4',
                                "Item Ledger Document Type"::"Sales Invoice",
                                "Item Ledger Document Type"::"Sales Shipment",
                                "Item Ledger Document Type"::"Sales Credit Memo",
                                "Item Ledger Document Type"::"Sales Return Receipt");

                            if ItemLedgerEntry.IsEmpty() then begin
                                // SI ESTÁ COMPLETAMENTE MUERTO HACE 3 AÑOS
                                vSinVenta36M := vCostoALaFecha;
                                vSinVenta12M := 0;
                            end else begin
                                // 2. POR DESCARTE: Si sí vendió en los últimos 3 años, verificamos si lleva al menos 12 meses congelado
                                FechaInicio12M := CalcDate('<-12M>', FechaFinalFiltro);
                                ItemLedgerEntry.SetRange("Posting Date", FechaInicio12M, FechaFinalFiltro);
                                if ItemLedgerEntry.IsEmpty() then begin
                                    // Si no vendió nada en el último año (pero sí entre el mes 13 y el 36)
                                    vSinVenta12M := vCostoALaFecha;
                                    vSinVenta36M := 0;
                                end else begin
                                    // Si tuvo ventas en el último año, NO es un producto inactivo.
                                    // Forzamos a que limpie ambas columnas para que no duplique.
                                    vSinVenta12M := 0;
                                    vSinVenta36M := 0;
                                end;
                            end;
                        end;
                    end;
                end else begin
                    vTotalStockCosto := vCostoALaFecha;
                    vSinVenta12M := 0;
                    vSinVenta36M := 0;


                    case true of
                        vIND = 0:
                            vN_D90 := vCostoALaFecha;
                        vIND >= 12:
                            vN_D30 := vCostoALaFecha;
                        vIND >= 6:
                            vN_D60 := vCostoALaFecha;
                        vIND >= 4:
                            vN_D90 := vCostoALaFecha;
                        vIND >= 2:
                            vN_D180 := vCostoALaFecha;
                        vIND >= 1:
                            vN_D360 := vCostoALaFecha;
                        vIND >= 0.6:
                            vN_1ANO_Mas := vCostoALaFecha;
                        vIND >= 0.2:
                            vN_2ANO_Mas := vCostoALaFecha;
                        else
                            vN_3ANO_Mas := vCostoALaFecha;
                    end;
                end;

                // Si el producto no tiene stock ni movimientos en el periodo, se salta
                if (vStockALaFecha = 0) and (vVentasPeriodo = 0) then
                    CurrReport.Skip();
            end;
        }
    }

    requestpage
    {
        layout
        {
            area(content)
            {
                group(Opciones)
                {
                    Caption = 'Filtros del Reporte';
                    field(FechaInicial; FechaInicialFiltro)
                    {
                        ApplicationArea = All;
                        Caption = 'Fecha Inicial';
                    }
                    field(FechaFinal; FechaFinalFiltro)
                    {
                        ApplicationArea = All;
                        Caption = 'Fecha Final';
                    }
                }
            }
        }
    }

    var
        FechaInicialFiltro: Date;
        FechaFinalFiltro: Date;
        vStockALaFecha: Decimal;
        vCostoALaFecha: Decimal;
        vVentasPeriodo: Decimal;
        ItemCategory: Record "Item Category";
        vNombreFamilia: Text;
        vCompras3M: Decimal;
        FechaInicio3M: Date;
        FechaInicio12M: Date;
        FechaInicio36M: Date;
        vIND: Decimal;
        // Variables de cubetas remanentes
        vN_D30: Decimal;
        vN_D60: Decimal;
        vN_D90: Decimal;
        vN_D180: Decimal;
        vN_D360: Decimal;
        vN_1ANO_Mas: Decimal;
        vN_2ANO_Mas: Decimal;
        vN_3ANO_Mas: Decimal;
        // Variables finales
        vSinVenta12M: Decimal;
        vSinVenta36M: Decimal;
        vTotalStockCosto: Decimal;

    local procedure ClearCalculations()
    begin
        vStockALaFecha := 0;
        vCostoALaFecha := 0;
        vVentasPeriodo := 0;
        vNombreFamilia := '';
        vCompras3M := 0;
        vIND := 0;
        vN_D30 := 0;
        vN_D60 := 0;
        vN_D90 := 0;
        vN_D180 := 0;
        vN_D360 := 0;
        vN_1ANO_Mas := 0;
        vN_2ANO_Mas := 0;
        vN_3ANO_Mas := 0;
        vSinVenta12M := 0;
        vSinVenta36M := 0;
        vTotalStockCosto := 0;
    end;
}

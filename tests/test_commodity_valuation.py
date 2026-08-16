from __future__ import annotations

from openpyxl import Workbook

from commodity_valuation import apply_commodity_normalization, is_commodity_producer
from advanced_analytics_v2 import _base_value


def _cvx_fixture():
    wb=Workbook(); wb.remove(wb.active)
    cd=wb.create_sheet('Company Data'); hf=wb.create_sheet('Historical Financials'); sc=wb.create_sheet('Three-Case Scenarios')
    cd['B4']='CVX'; cd['B6']='Energy'; cd['B7']='Oil & Gas Integrated'; cd['B8']=197.70; cd['B9']=1.961603364; cd['B10']=387.809; cd['B12']=8.530; cd['B13']=37.075; cd['B14']=28.545
    years=[2020,2021,2022,2023,2024,2025]
    rev=[94.471,155.606,235.717,196.913,193.414,184.432]
    op=[-6.942,16.104,39.950,26.229,18.917,16.674]
    ocf=[10.577,29.187,49.618,35.609,31.493,33.939]
    cap=[8.922,8.654,12.314,15.829,16.375,17.347]
    da=[17.181,17.251,18.990,19.207,20.127,20.132]
    for i,y in enumerate(years,2):
        hf.cell(3,i,y); hf.cell(4,i,rev[i-2]); hf.cell(9,i,op[i-2]); hf.cell(14,i,ocf[i-2]); hf.cell(15,i,cap[i-2]); hf.cell(18,i,da[i-2]); hf.cell(21,i,0)
    # Deliberately reproduce the old over-aggressive generic scenario.
    generic_growth=[.2384,-.0848,.1236,.1138,.1040,.0942,.0844,.0746,.0648,.0550]
    for block in (range(2,12),range(14,24),range(26,36)):
        for c,g in zip(block,generic_growth):
            sc.cell(12,c,g); sc.cell(14,c,.22); sc.cell(18,c,.10916); sc.cell(20,c,.09406)
    sc['C6']=.06426; sc['C7']=.03; sc['C8']=.21; sc['AL4']=-.0038
    return wb


def test_cvx_normalizes_peak_cycle_dcf():
    wb=_cvx_fixture()
    assert is_commodity_producer(wb,'CVX')
    before=_base_value(wb)
    meta=apply_commodity_normalization(wb,'CVX',{})
    after=_base_value(wb)
    sc=wb['Three-Case Scenarios']
    assert meta['applied'] is True
    assert 'Commodity Valuation' in wb.sheetnames
    assert abs(sc['C7'].value-.02)<1e-12
    assert sc.cell(12,23).value<=.03
    assert sc.cell(14,23).value<.18
    assert 17.5<=meta['base_capex_nominal'][0]<=20.0
    assert before>500, before
    assert 150<after<400, after
    assert after<before*.65, (before,after)


def test_noncommodity_company_is_untouched():
    wb=_cvx_fixture(); wb['Company Data']['B4']='GOOGL'; wb['Company Data']['B6']='Communication Services'; wb['Company Data']['B7']='Internet Content & Information'
    original=wb['Three-Case Scenarios']['C7'].value
    meta=apply_commodity_normalization(wb,'GOOGL',{})
    assert meta['applied'] is False
    assert wb['Three-Case Scenarios']['C7'].value==original
    assert 'Commodity Valuation' not in wb.sheetnames


if __name__=='__main__':
    test_cvx_normalizes_peak_cycle_dcf(); test_noncommodity_company_is_untouched(); print('commodity valuation tests passed')

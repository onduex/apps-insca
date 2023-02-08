# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import xlwings as xw
import zeep


@xw.sub
def main():
    url = 'http://srwinvlt01/AutodeskDM/Services/Filestore/v28/AuthService.svc?wsdl'
    client = zeep.Client(wsdl=url)
    sign_in = client.service.SignIn(userName='mecanic', userPassword='mecanic', knowledgeVault='tavil')
    token = sign_in.Authorization
    user_id = sign_in.User.Id
    wb = xw.Book.caller()
    numbers = wb.sheets[0].range('C2').expand('down')
    i = 1  # Indice para recorrer columnas, uno menos que el range de numbers

    # Acceso al servicio de Artículos
    url_item = 'http://srwinvlt01/AutodeskDM/Services/v28/ItemService.svc?wsdl'
    settings = zeep.Settings(extra_http_headers={'Authorization': token})
    client_item = zeep.Client(wsdl=url_item, settings=settings)

    # Acceso al servicio de Propiedades
    url_prop = 'http://srwinvlt01/AutodeskDM/Services/v28/PropertyService.svc?wsdl'
    client_prop = zeep.Client(wsdl=url_prop, settings=settings)
    for rec in numbers:
        i = i + 1
        GetVaultItems = client_item.service.FindItemRevisionsBySearchConditions(
            searchConditions={'SrchCond': {'PropDefId': 31,
                                    'SrchOper': 3,
                                    'PropTyp': 'SingleProperty',
                                    'SrchRule': 'Must',
                                    'SrchTxt': rec.formula,
                                    }},
            bRequestLatestOnly=True,
        )
        if GetVaultItems.searchstatus.TotalHits == 1:
            Title = GetVaultItems.FindItemRevisionsBySearchConditionsResult.Item[0].Title
            Id = GetVaultItems.FindItemRevisionsBySearchConditionsResult.Item[0].Id
            GetVaultItemsCustomProperties = client_prop.service.GetProperties(
                entityClassId='ITEM',
                entityIds={'long': Id},
                propertyDefIds={'long': [37, 141, 165, 171, 172]},
            )
            Units = GetVaultItemsCustomProperties[0].Val
            SapMatGr = GetVaultItemsCustomProperties[1].Val
            SapMatTy = GetVaultItemsCustomProperties[2].Val
            Aprov = GetVaultItemsCustomProperties[3].Val
            AprovEsp = GetVaultItemsCustomProperties[4].Val
        else:
            Title = '####'
            SapMatGr = '####'
            SapMatTy = '####'
            Aprov = '####'
            AprovEsp = '####'
            Units = '####'
        if not wb.sheets[0].range('D' + str(i)).value:
            wb.sheets[0].range('D' + str(i)).value = Title
        if not wb.sheets[0].range('E' + str(i)).value:
            wb.sheets[0].range('E' + str(i)).value = SapMatGr
        if not wb.sheets[0].range('F' + str(i)).value:
            wb.sheets[0].range('F' + str(i)).value = SapMatTy
        if not wb.sheets[0].range('G' + str(i)).value:
            wb.sheets[0].range('G' + str(i)).value = Aprov
        if not wb.sheets[0].range('H' + str(i)).value:
            wb.sheets[0].range('H' + str(i)).value = AprovEsp
        if not wb.sheets[0].range('I' + str(i)).value:
            wb.sheets[0].range('I' + str(i)).value = Units


if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()

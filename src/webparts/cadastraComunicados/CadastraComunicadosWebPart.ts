//necessario para carregar componentes externos, nesse caso iremos carregar o bootstrap
import { SPComponentLoader } from '@microsoft/sp-loader';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CadastraComunicadosWebPart.module.scss';
import * as strings from 'CadastraComunicadosWebPartStrings';

//carrega o pnp
import { sp, Item, ItemAddResult } from '@pnp/sp';

//carrega bootstrap
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');

export interface IListItem {  
  Id: number;  
  Title?: string;  
  Link?: string;
} 

export interface ICadastraComunicadosWebPartProps {
  description: string;
}

export default class CadastraComunicadosWebPart extends BaseClientSideWebPart<ICadastraComunicadosWebPartProps> {

  public onInit(): Promise<void> 
  {
      return super.onInit().then(_ => {
          sp.setup({
          spfxContext: this.context
          });
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.cadastraComunicados }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">              
              <div class="row">
              <div class="col-md-8">
                <h2>Criar comunicado</h2>
                <div class="form-group">
                  <input type="text" id="txtTitulo" placeholder="Título do comunicado" class="form-control"/>
                  <input type="text" id="txtLink" placeholder="https://www.google.com.br" class="form-control"/>
                  <br/>
                  <button type="button" class="btn btn-success criarComunicado">Salvar</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;

      this.setButtonsEventHandlers();
  }

  private setButtonsEventHandlers(): void {
    const webPart: CadastraComunicadosWebPart = this;
    this.domElement.querySelector('button.criarComunicado').addEventListener('click', () => { webPart.criarComunicado(); });
  }

  private criarComunicado(): void 
  {
   
    sp.web.lists.getByTitle(this.properties.description).items.add({  
      'Title': document.getElementById('txtTitulo')["value"],
      'Link': document.getElementById('txtLink')["value"]

    }).then((result: ItemAddResult): void => {  
      const item: IListItem = result.data as IListItem;  
      console.log(`Id do item criado ${item.Id}`);
      window.location.href = this.context.pageContext.web.absoluteUrl + "/SitePages/ListarComunicados.aspx";
    }, (error: any): void => {  
      console.log('Erro ao cadastrar o comunicado: ' + error);  
    });

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class PCFComentarioPortal implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private divelementForm: HTMLDivElement;
	private divalertaMensagem: HTMLDivElement;
	private pMensagem: HTMLElement;
	private entityId: string;
	private contactId: string;
	private saveButtonInnerText: string;
	private clearButtonInnerText: string;
	private textAreaInnerText: string;
	private SucessoInnerText: string;
	private ErrorInnerText: string;

	private localContext: ComponentFramework.Context<IInputs>;
	private webApiContext: ComponentFramework.WebApi;
	private localNotifyOutputChanged: () => void;
	private localContainer: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this.webApiContext = context.webAPI;
		this.localContext = context;
		this.localContainer = container;
		this.localContainer.style.maxHeight = "400px";
		this.localContainer.style.minHeight = "200px";
		this.localContainer.style.textAlign = "center";
		this.divelementForm = document.createElement("div");
		this.setTextPorIdioma();
		this.criaDivAlerta();
		this.createForm(this.divelementForm);
		this.localContainer.appendChild(this.divalertaMensagem);
		this.localContainer.appendChild(this.divelementForm);
		this.entityId = (<any>context).page.entityId;
		console.log(Xrm.Page.getControl<Xrm.Page.LookupControl>("primarycontactid").getAttribute().getValue());
	}
	public setTextPorIdioma(): void {
		let idioma: number = Xrm.Page.context.getUserLcid();
		switch (idioma) {
			case 1046:
				this.saveButtonInnerText = "Enviar";
				this.clearButtonInnerText = "Limpar";
				this.textAreaInnerText = "Digite o texto";
				this.SucessoInnerText = "Sucesso";
				this.ErrorInnerText = "Erro ao criar comentário, notifique o administradoe";
				break;
			case 3082:
				this.saveButtonInnerText = "Enviar";
				this.clearButtonInnerText = "Limpar";
				this.textAreaInnerText = "Insertar texto";
				this.SucessoInnerText = "Sucesso";
				this.ErrorInnerText = "Error al crear el comentario, notifique al administrador";
				break;
			default:
				this.saveButtonInnerText = "Submit";
				this.clearButtonInnerText = "Cancel";
				this.textAreaInnerText = "Enter text";
				this.SucessoInnerText = "success";
				this.ErrorInnerText = "Error creating comment, please notify the Administrator";
				break;
		}
	}
	public createForm(_divElementForm: HTMLDivElement): void {

		let divTextArea: HTMLDivElement = this.createDiv("itsm_item itsm_width_100");
		let textArea: HTMLTextAreaElement = document.createElement("textarea");
		textArea.className = "itsm_textarea ";
		textArea.placeholder = this.textAreaInnerText ;
		textArea.maxLength = 100000;
		divTextArea.append(textArea);

		let divInput: HTMLDivElement = this.createDiv("itsm_item");
		let input: HTMLInputElement = document.createElement("input");
		input.type = "file";
		input.title = "arquivo";
		input.className = "itsm_input";
		input.style.opacity = "100";
		input.style.position = "relative";
		input.style.pointerEvents = "all";
		input.style.cursor = "pointer";
		input.style.width = "100%";
		input.style.height = "30%";

		divInput.append(input);

		let divButtons: HTMLDivElement = this.createDiv("itsm_item");
		let buttonSave: HTMLButtonElement = this.createButton("itsm_button_comment itsm_save");
		buttonSave.innerHTML = "<p> " + this.saveButtonInnerText +" </p>";
		buttonSave.title = this.saveButtonInnerText;
		let buttonCancel: HTMLButtonElement = this.createButton("itsm_button_comment itsm_cancel");
		buttonCancel.innerHTML = "<p> " +this.clearButtonInnerText+ " </p>";
		buttonCancel.title = this.clearButtonInnerText;
		buttonSave.addEventListener("click", this.chamaCriacaoRegistros.bind(this, buttonSave, textArea, input));
		buttonCancel.addEventListener("click", this.limpaCampos.bind(this, textArea, input));
		divButtons.append(buttonSave);
		divButtons.appendChild(buttonCancel);

		_divElementForm.append(divTextArea);
		_divElementForm.appendChild(divInput);
		_divElementForm.appendChild(divButtons);
	}
	public createDiv(classe: string): HTMLDivElement {
		let div: HTMLDivElement;
		div = document.createElement("div");
		div.className = classe;
		return div;
	}
	public createButton(classe: string): HTMLButtonElement {
		let button: HTMLButtonElement;
		button = document.createElement("button");
		button.className = classe;
		return button;
	}
	private async chamaCriacaoRegistros(botao: HTMLButtonElement, textArea: HTMLTextAreaElement, inputFiles: HTMLInputElement): Promise<void> {
		this.desabilitaHabilitaBotao(botao, true);
		let descricao: string = textArea.value;
		let idComentario: string;
		let IdAnotacao: string = "";
		if (descricao != null && descricao.replace(" ", "").length > 0) {
			idComentario = await this.criaComentario(descricao);
			//idComentario = "await this.criaComentario(descricao)";
			if (idComentario != "") {
				if (inputFiles != null && inputFiles.files?.length) {
					for (let index = 0; index < inputFiles.files?.length; index++) {
						let base: any = await this.getBase64(inputFiles.files[index]);
						IdAnotacao = await this.createAnotacao(idComentario, "adx_portalcomment", base, inputFiles.files[index]);
					}
				}
				this.mostraMensagenAlerta(this.SucessoInnerText, "rgb(0, 255, 0,0.2)");
				this.desabilitaHabilitaBotao(botao, false);
				this.limpaCampos(textArea, inputFiles);

				//this.limpaCampos(textArea, inputElement);
			} else {
				this.mostraMensagenAlerta(this.ErrorInnerText, "rgb(255, 0, 0,0.2)");
				this.desabilitaHabilitaBotao(botao, false);
			}
		} else {
			this.mostraMensagenAlerta("O campo de descrição é obrigatório", "rgb(255, 0, 0,0.2)");
			this.desabilitaHabilitaBotao(botao, false);
		}
	}
	private desabilitaHabilitaBotao(botao: HTMLButtonElement, bool: boolean) {
		botao.disabled = bool;
	}
	private mostraMensagenAlerta(mensgem: string, background: string): void {
		this.divalertaMensagem.style.display = "block";
		this.divalertaMensagem.style.backgroundColor = background;
		this.divalertaMensagem.innerText = mensgem;
		setTimeout(() => {
			this.divalertaMensagem.style.display = "none";
		}, 2000);
	}

	private criaDivAlerta(): void {
		let divAlerta: HTMLDivElement = this.createDiv("itsm_alerta");
		let _pMensagem: HTMLElement = document.createElement("p");
		_pMensagem.className = "itsm_message_alert";
		this.pMensagem = _pMensagem;
		//divAlerta.append(this.pMensagem);
		this.divalertaMensagem = divAlerta;
	}

	private limpaCampos(textArea: HTMLTextAreaElement, inputElement: HTMLInputElement): any {
		textArea.value = "";
		if (inputElement != null && inputElement.files != undefined && inputElement.files?.length > 0) {
			inputElement.value = "";
		}

	}

	private async criaComentario(descricao: string): Promise<string> {
		let idContact: string = Xrm.Page.getAttribute<Xrm.Page.LookupAttribute>("primarycontactid").getValue()[0].id;
		let idSystemUser: string = Xrm.Page.context.getUserId();
		var entity = {
			"description": descricao,
			["regardingobjectid_incident@odata.bind"]: "/incidents(" + this.entityId.replace("{", "").replace("}", "") + ")",
			"adx_portalcommentdirectioncode": 2,
			"statuscode": 1,
			["adx_portalcomment_activity_parties"]: [
				{
					"partyid_systemuser@odata.bind": "/systemusers(" + idSystemUser.replace("{", "").replace("}", "") + ")",
					"participationtypemask": 1  ///From Email
				},
				{
					"partyid_contact@odata.bind": "/contacts(" + idContact.replace("{", "").replace("}", "") + ")",
					"participationtypemask": 2  ///To Email
				}]
		};
		console.log(entity);
		var newEntityId: string = "";
		await this.webApiContext.createRecord("adx_portalcomment", entity).then(
			function success(result) {
				newEntityId = "" + result.id;
			},
			function (error) {
				console.log(error.message);
			}
		);
		return newEntityId;
	}

	private async createAnotacao(Id: string, entityLogicalName: string, base: string, file: any): Promise<string> {
		let criado: string = "";
		console.log("mimetype:");
		console.log(file);
		var attachment = {
			"subject": "Anexo: " + file.name,
			"filename": file.name,
			"filesize": file.size,
			"mimetype": file.type,
			"objecttypecode": entityLogicalName,
			"documentbody": base,
			[`objectid_${entityLogicalName}@odata.bind`]: `/${entityLogicalName}s(${Id})`
		};
		console.log(attachment);
		await this.webApiContext.createRecord("annotation", attachment).then(
			function success(result) {
				criado = "" + result.id;
				console.log("Arquivo criado");
				return criado;
			},
			function (error) {
				console.log(error.message);
				return criado;
			}
		);
		return criado;
	}

	private getBase64(file: File): Promise<string> {
		return new Promise((resolve, reject) => {
			const reader = new FileReader();
			reader.onload = (f) => resolve((<string>reader.result).split(',')[1]);
			reader.onerror = error => reject(error);
			reader.readAsDataURL(file);
		});
	}
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}
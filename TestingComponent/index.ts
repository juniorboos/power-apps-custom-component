/* eslint-disable no-undef */
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ResultReason } from "microsoft-cognitiveservices-speech-sdk";
import axios from "axios";

import * as speechsdk from "microsoft-cognitiveservices-speech-sdk";

export class TestingComponent
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  /**
   * Empty constructor.
   */
  constructor() {}

  private _container: HTMLDivElement;
  private _context: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged: () => void;
  private _refreshData: EventListenerOrEventListenerObject;
  private labelElement: HTMLLabelElement;
  private displayText: string;

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    // Add control initialization code
    this._context = context;
    this._container = document.createElement("div");
    this._notifyOutputChanged = notifyOutputChanged;
    this._refreshData = this.refreshData.bind(this);

    // creating a HTML label element that shows the value that is set on the linear range control
    this.labelElement = document.createElement("label");
    this.labelElement.setAttribute("class", "LinearRangeLabel");
    this.labelElement.setAttribute("id", "lrclabel");

    // retrieving the latest value from the control and setting it to the HTMl elements.
    this.displayText = context.parameters.sampleProperty.raw!;
    this.labelElement.innerHTML = context.parameters.sampleProperty.formatted
      ? context.parameters.sampleProperty.formatted
      : "0";

    // appending the HTML elements to the control's HTML container element.
    this._container.appendChild(this.labelElement);
    container.appendChild(this._container);

    this.sttFromMic();
  }

  public async getTokenOrRefresh() {
    const speechKey = process.env.SPEECH_KEY;
    const speechRegion = "westeurope";

    const headers = {
      headers: {
        "Ocp-Apim-Subscription-Key": speechKey,
        "Content-Type": "application/x-www-form-urlencoded",
      },
    };

    try {
      const tokenResponse = await axios.post(
        `https://${speechRegion}.api.cognitive.microsoft.com/sts/v1.0/issueToken`,
        null,
        headers
      );
      console.log("Token fetched from back-end: " + tokenResponse.data);
      return { authToken: tokenResponse.data, region: speechRegion };
    } catch (err) {
      return { authToken: null, error: err.response.data };
    }
  }

  public async sttFromMic() {
    const tokenObj = await this.getTokenOrRefresh();
    const speechConfig = speechsdk.SpeechConfig.fromAuthorizationToken(
      tokenObj.authToken,
      tokenObj.region || ""
    );
    speechConfig.speechRecognitionLanguage = "en-US";

    const audioConfig = speechsdk.AudioConfig.fromDefaultMicrophoneInput();
    const recognizer = new speechsdk.SpeechRecognizer(
      speechConfig,
      audioConfig
    );

    this.displayText = "speak into your microphone...";

    recognizer.recognizeOnceAsync((result) => {
      let displayText;
      if (result.reason === ResultReason.RecognizedSpeech) {
        displayText = `RECOGNIZED: Text=${result.text}`;
      } else {
        displayText =
          "ERROR: Speech was cancelled or could not be recognized. Ensure your microphone is working properly.";
      }

      this.displayText = displayText;
    });
  }

  public refreshData(evt: Event): void {
    this.labelElement.innerHTML = this.displayText;
    this._notifyOutputChanged();
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
    this.displayText = context.parameters.sampleProperty.raw!;
    this._context = context;
    this.labelElement.innerHTML = context.parameters.sampleProperty.formatted
      ? context.parameters.sampleProperty.formatted
      : "";
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      sampleProperty: this.displayText,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}

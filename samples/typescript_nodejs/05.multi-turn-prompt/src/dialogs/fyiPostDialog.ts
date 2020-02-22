// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
import { StatePropertyAccessor, TurnContext, UserState } from "botbuilder";
import {
  ChoiceFactory,
  ChoicePrompt,
  ComponentDialog,
  ConfirmPrompt,
  DialogSet,
  DialogTurnStatus,
  PromptValidatorContext,
  TextPrompt,
  WaterfallDialog,
  WaterfallStepContext
} from "botbuilder-dialogs";
import { FyiPost } from "../fyiPost";

const SOURCETYPE_PROMPT = "SOURCETYPE_PROMPT";
const CONFIRM_PROMPT = "CONFIRM_PROMPT";
const DESCRIPTION_PROMPT = "DESCRIPTION_PROMPT";
const URL_PROMPT = "URL_PROMPT";
const PRIORITY_PROMPT = "PRIORITY_PROMPT";
const USER_PROFILE = "USER_PROFILE";
const WATERFALL_DIALOG = "WATERFALL_DIALOG";

export class FyiPostDialog extends ComponentDialog {
  private fyiPost: StatePropertyAccessor<FyiPost>;

  constructor(userState: UserState) {
    super("fyiPostDialog");

    this.fyiPost = userState.createProperty(USER_PROFILE);

    this.addDialog(new TextPrompt(DESCRIPTION_PROMPT));
    this.addDialog(new TextPrompt(URL_PROMPT));
    this.addDialog(new ChoicePrompt(SOURCETYPE_PROMPT));
    this.addDialog(new ChoicePrompt(PRIORITY_PROMPT));
    this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));

    this.addDialog(
      new WaterfallDialog(WATERFALL_DIALOG, [
        this.sourceTypeStep.bind(this),
        this.urlStep.bind(this),
        this.descriptionStep.bind(this),
        this.priorityStep.bind(this),
        this.confirmStep.bind(this),
        this.summaryStep.bind(this)
      ])
    );

    this.initialDialogId = WATERFALL_DIALOG;
  }

  /**
   * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
   * If no dialog is active, it will start the default dialog.
   * @param {*} turnContext
   * @param {*} accessor
   */
  public async run(turnContext: TurnContext, accessor: StatePropertyAccessor) {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);

    const dialogContext = await dialogSet.createContext(turnContext);
    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    }
  }

  private async sourceTypeStep(stepContext: WaterfallStepContext<FyiPost>) {
    // WaterfallStep always finishes with the end of the Waterfall or with another dialog; here it is a Prompt Dialog.
    // Running a prompt here means the next WaterfallStep will be run when the users response is received.
    return await stepContext.prompt(SOURCETYPE_PROMPT, {
      choices: ChoiceFactory.toChoices([
        "Website",
        "Konferenz",
        "Literatur",
        "Sonstiges"
      ]),
      prompt: "Bitte sage mir kurz, um was für eine Quelle es sich handelt"
    });
  }

  private async urlStep(stepContext: WaterfallStepContext<FyiPost>) {
    stepContext.options.sourceType = stepContext.result.value;
    return await stepContext.prompt(
      URL_PROMPT,
      "Bitte gebe mir die URL des Webinhaltes."
    );
  }

  private async descriptionStep(stepContext: WaterfallStepContext<FyiPost>) {
    stepContext.options.url = stepContext.result;
    return await stepContext.prompt(
      DESCRIPTION_PROMPT,
      "Was möchtest du den anderen hinsichtlich der Quelle sagen (diese Angabe wird 1 zu 1 in übernommen)?"
    );
  }

  private async priorityStep(stepContext: WaterfallStepContext<FyiPost>) {
    stepContext.options.description = stepContext.result;
    return await stepContext.prompt(PRIORITY_PROMPT, {
      choices: ChoiceFactory.toChoices([
        "Eher unwichtig",
        "Wichtig",
        "Dringend"
      ]),
      prompt: "Bitte sage mir kurz, wie wichtig diese Quelle ist."
    });
  }

  private async confirmStep(stepContext: WaterfallStepContext<FyiPost>) {
    // WaterfallStep always finishes with the end of the Waterfall or with another dialog, here it is a Prompt Dialog.
    stepContext.options.priority = stepContext.result.value;
    return await stepContext.prompt(CONFIRM_PROMPT, {
      prompt: "Sind deine Angaben so in Ordnung?"
    });
  }

  private async summaryStep(stepContext: WaterfallStepContext<FyiPost>) {
    if (stepContext.result) {
      // Get the current profile object from user state.
      const fyiPost = await this.fyiPost.get(
        stepContext.context,
        new FyiPost()
      );
      const stepContextOptions = stepContext.options;
      fyiPost.sourceType = stepContextOptions.sourceType;
      fyiPost.url = stepContextOptions.url;
      fyiPost.description = stepContextOptions.description;
      fyiPost.priority = stepContextOptions.priority;

      let msg = `Ich hab deine Empfehlung vom Typ *${fyiPost.sourceType}* wie folgt abgespeichert:\n\n`;
      msg += `**URL:** ${fyiPost.url}. \n \n`;
      msg += `\n`;
      msg += ` **Beschreibung:** ${fyiPost.description}.\n`;
      msg += `\n`;
      msg += ` Ebenfalls habe ich die Quelle als *${fyiPost.priority}* einsortiert.`;

      await stepContext.context.sendActivity(msg);
    } else {
      await stepContext.context.sendActivity(
        "Die Daten wurden nicht gespeichert."
      );
    }

    // WaterfallStep always finishes with the end of the Waterfall or with another dialog, here it is the end.
    return await stepContext.endDialog();
  }

  private async descriptionPromptValidator(
    promptContext: PromptValidatorContext<Text>
  ) {
    // This condition is our validation rule. You can also change the value at this point.
    return promptContext.recognized.succeeded;
  }
}

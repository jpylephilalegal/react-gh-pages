import * as React from "react";
import * as ReactDOM from "react-dom";

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, ActionButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { IComboBox } from 'office-ui-fabric-react/lib/components/ComboBox/ComboBox.types';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons();

const Office = window.Office;
const Word = window.Word;

declare var Cookies:any;
declare var window: any;
declare module JSX {
    interface IntrinsicElements {
        "AddinApp": AddinApp
    }
}

class AddinApp extends React.Component<any, any> {
    constructor(props: {}) {
        super(props);
        var serverName = Cookies.get('serverName');
        this.state = {
            showServerName: serverName ? false : true,
            showServerNameError: false,
            frameStage: serverName ? "wait" : "standby",
            serverName: serverName,
            showApp: false,
            interviewList: [],
            varList: [],
            varOptions: [],
            vocabList: [],
            currentInterview: null,
            interviewOptions: [],
            selectedVar: null,
            findReplace: false,
            applyFormat: ""
        }
        this.receiveMessage = this.receiveMessage.bind(this);
        this.handleServerNameChange = this.handleServerNameChange.bind(this);
        this.fetchVars = this.fetchVars.bind(this);
        this.fetchFiles = this.fetchFiles.bind(this);
        this.testServer = this.testServer.bind(this);
        this.handleSetServer = this.handleSetServer.bind(this);
        this.handleInterviewChange = this.handleInterviewChange.bind(this);
        this.varChanged = this.varChanged.bind(this);
        this.handleFindReplaceChange = this.handleFindReplaceChange.bind(this);
        this.handleApplyFormatChanged = this.handleApplyFormatChanged.bind(this);
    }

    render() {
        return (
            <div>
                <div className={ this.state.showServerName ? 'shownelement' : 'hiddenelement' }>
                    <TextField
                        label='Your docassemble server'
                        onChanged={ this.handleServerNameChange }
                    />
                    <div className={ this.state.showServerNameError ? 'errormessage' : 'hiddenelement' }>
                        Please enter a valid URL.
                    </div>
                    <DefaultButton
                        text='Connect to server'
                        onClick={ this.handleSetServer }
                    />
                </div>
                <iframe id="server" src={ this.state.serverName ? this.state.serverName + '/officeaddin' : 'about:blank'} className={ this.state.frameStage == "wait" ? 'shownelement' : 'hiddenelement' } />
                <main id="app-body" className={ this.state.showApp ? 'ms-welcome__main' : 'hiddenelement' }>
                    <Dropdown
                        label='Interview'
                        onChanged={ this.handleInterviewChange }
                        options={ this.state.interviewOptions }

                    />
                    <h2 className="ms-font-xl">Insert Variables</h2>
                    <ComboBox
                        label="Variable Name"
                        allowFreeform={true}
                        autoComplete='on'
                        onChanged={ this.varChanged }
                        options={ this.state.varOptions }
                    />
                    <Checkbox
                        label="Find and Replace All"
                        onChange={ this.handleFindReplaceChange }
                    />
                    <Dropdown
                        label="Apply Format"
                        onChanged={ this.handleApplyFormatChanged }
                        options={ this.applyFormatOptions }
                        defaultSelectedKey=""
                    />
                    <DefaultButton
                        iconProps={ { iconName: 'CirclePlus' } }
                        onMenuClick={ this.insertVariable }
                        text="Insert Variable"
                    />
                </main>
            </div>);
    }
    
    fetchVars(yamlFile: string): void {
        if (yamlFile == null){
            console.log("fetchVars: yamlFile was null");
            return;
        }
        var server : any = document.getElementById('server');
        var action = Object();
        action.action = "fetchVars";
        action.file = yamlFile;
        server.contentWindow.postMessage(action, this.state.serverName);
    }

    testServer() {
        var server : any = document.getElementById('server');
        var action = Object();
        action.action = "test";
        server.contentWindow.postMessage(action, this.state.serverName);
    }

    fetchFiles() {
        var server : any = document.getElementById('server');
        var action = Object();
        action.action = "fetchFiles";
        server.contentWindow.postMessage(action, this.state.serverName);
    }
    
    handleServerNameChange(newvalue: any) {
        this.setState({ serverName: newvalue });
    }

    handleFindReplaceChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
        this.setState({ findReplace: isChecked });
    }

    private applyFormatOptions =
        [
            {key: "", text: "(none)"},
            {key: "currency", text: "Currency (localized)"},
            {key: "capitalize", text: "Capitalize"},
            {key: "indefinite_article", text: "Indefinite Article"},
            {key: "nice_number", text: "Nice Number"},
            {key: "noun_plural", text: "Pluralize noun"},
            {key: "noun_singular", text: "Singularize noun"},
            {key: "ordinal_number", text: "Ordinal Number"},
            {key: "title_case", text: "Title Case"},
            {key: "verb_past", text: "Past tense verb"},
            {key: "verb_present", text: "Present tense verb"},
            {key: "fix_punctuation", text: "Fix punctuation"},
        ];

    handleApplyFormatChanged(newoption: any) {
        this.setState({ applyFormat: newoption.key });
    }

    handleSetServer(event: any) {
        console.log("handleSetServer");
        if (!validateUrl(this.state.serverName)) {
            this.setState({showServerNameError: true});
            return;
        }
        else {
            this.setState({showServerNameError: false});
        }
        Cookies.set('serverName', this.state.serverName, { expires: 999999 });
        this.setState({frameStage: 'wait', showServerName: false});
    }

    varChanged = (option: IComboBoxOption, index: number, value: string): void => {
        console.log('varChanged() is called: option = ' + JSON.stringify(option));
        if (option !== undefined) {
            this.setState({
                selectedVar: option.key
            });
        } else if (index !== undefined && index >= 0 && index < this.state.varOptions.length) {
            this.setState({
                selectedVar: this.state.varOptions[index].key
            });
        } else if (value !== undefined) {
            const newOption: IComboBoxOption = { key: value, text: value };
            this.setState({
                varOptions: [...this.state.varOptions, newOption],
                selectedVar: newOption.key
            });
        }
    }

    handleInterviewChange(newoption: any) {
        console.log("handleInterviewChange");
        if (newoption.key){
            this.setState({currentInterview: newoption.key});
            this.fetchVars(newoption.key);
        }
        else{
            console.log("handleInterviewChange: no new option");
        }
    }

    receiveMessage(event: any) {
        console.log("receiveMessage");
        //console.log("receiveMessage " + JSON.stringify(event.data));
        if (!this.state) {
            return;
        }
        if (event.origin !== this.state.serverName) {
            console.log("Message received from improper origin " + event.origin);
            return;
        }
        console.log("Received action " + event.data.action);
        if (event.data.action == 'initialize') {
            this.setState({showApp: true, frameStage: 'loaded'});
            this.fetchFiles();
        }
        if (event.data.action == 'files') {
            var arr = Array();
            var n = event.data.files.length;
            for (var i = 0; i < n; i++) {
                arr.push({key: event.data.files[i], text: event.data.files[i]});
            }
            this.setState({interviewList: event.data.files, interviewOptions: arr});
        }
        if (event.data.action == 'vars') {
            var arr = Array();
            var n = event.data.vars.var_list.length;
            for(var i = 0; i < n; ++i){
                var info = event.data.vars.var_list[i];
                if (!info.hide){
                    var newOption: IComboBoxOption = {key: info.var, text: info.var}
                    arr.push(newOption);
                }
            }
            console.log("varOptions is " + arr.length + " elements long from " + n);
            console.log("varOptions is " + JSON.stringify(arr));
            this.setState({varOptions: arr, varList: event.data.vars, vocabList: event.data.vocab});
            this.fetchFiles();
        }
    }
    
    componentDidMount() {
        console.log("adding listener");
        window.addEventListener("message", this.receiveMessage, false);
    }

    insertVariable() {
        Word.run(async (context: any) => {
            const range = context.document.getSelection();

            var variableName = this.state.selectedVar;

            if (variableName == null){
                return;
            }
            var variableReplaceAll = this.state.findReplace;
            var variableFormat = this.state.applyFormat;

            var textToInsert: string;
            if (variableFormat == "") {
                textToInsert = variableName;
            } else {
                textToInsert = variableFormat + '(' + variableName + ')';
            }

            range.load('text');

            if (! variableReplaceAll) {
                range.insertText('{{ ' + variableName + ' }}','Replace');
            } else {
                await context.sync();
                var textToReplace = range.text;

                // FIXME: We need to ignore Jinja statements and expressions -- search looks inside them now
                var results = context.document.body.search(textToReplace.trim(), {matchWholeWord: true}); // Word Online seems to select spaces next to a word you double-click on
                context.load(results);
                
                await context.sync();

                for (var i = 0; i < results.items.length; i++) {
                    results.items[i].insertText('{{ ' + textToInsert + ' }}', "Replace");
                }
            }

            await context.sync();
        });
    }

    /* ifPara() {
       Word.run(async (context: any) => {
     *         const range = context.document.getSelection();
     *         var ifExpression = document.getElementById('inputIfExpression').value;

     *         // Read the range text
     *         range.load('text');
     *         var textBefore = '{%p if ' + ifExpression + ' %}';

     *         range.insertParagraph(textBefore,'Before');
     *         range.insertParagraph('{%p endif %}','After');

     *         await context.sync();
     *         console.log(`The selected text was ${range.text}.`);
       });
     * }

     * ifInline() {
       Word.run(async (context: any) => {
     *         const range = context.document.getSelection();
     *         var ifExpression = document.getElementById('inputIfExpression').value;
     *         var textBefore = '{% if ' + ifExpression + ' %}';

     *         // Read the range text
     *         range.load('text');

     *         range.insertText(textBefore,'Before');
     *         range.insertText('{% endif %}','After');

     *         await context.sync();
     *         console.log(`The selected text was ${range.text}.`);
       });
     * }

     * listPara() {
       Word.run(async (context: any) => {
     *         const range = context.document.getSelection();
     *         var listVariableName = document.getElementById('inputListVariableName').value;
     *         var onlyTrue = document.getElementById('checkboxOnlyTrue').checked;
     *         if (onlyTrue) {
       var textBefore = '{%p for item in ' + listVariableName + '.true_values() %}'; 
     *         } else {
       var textBefore = '{%p for item in ' + listVariableName + '%}'; 
     *         }
     *         // Read the range text
     *         range.load('text');
     *         range.insertText('{{ item }}','Replace');
     *         range.insertParagraph(textBefore,'Before');
     *         range.insertParagraph('{%p endfor %}','After');

     *         await context.sync();
     *         console.log(`The selected text was ${range.text}.`);
       });
     * }

     * commentPara() {
       Word.run(async (context: any) => {
     *         const range = context.document.getSelection();
     *         
     *         // Read the range text
     *         range.load('text');
     *         await context.sync(); // Guess this has a performance penalty?
     *         
     *         // Regexp with 3 groups: {# , text between comments, #}. We match both whitespace and non-whitespace, including newlines
     *         var re = new RegExp('({#)([\\s\\S]*)(#})');
     *         var matches = re.exec(range.text);

     *         if (matches) { // index 1 is the uncommented string
       // This is not correct as it removes formatting from the text
       // This sample looks like it shows how to do it correctly: https://github.com/OfficeDev/Word-Add-in-JS-SpecKit/blob/master/scripts/boilerplate.js in addBoilerplateParagraph
       // we should use var paragraphs = context.document.getSelection().paragraphs; and then loop through paragraph collection
       range.insertText(matches[2],'Replace'); 
       console.log('Removed comments.')
     *         } else {
       range.insertParagraph('{#','Before');
       range.insertParagraph('#}','After');
       console.log('Added comments.')
       // we should extend the selection to include the newly added text
     *         }
     *         await context.sync();
       });
     * }

     * insertTemplate() {
       Word.run(async (context: any) => {
     *         const range = context.document.getSelection();
     *         var templateName : any = document.getElementById('inputTemplateName').value;
       
     *         var templateOptions : any = document.getElementById('inputTemplateOptions').value;
     *         if (templateOptions == "") {
       var textBefore = '{{p include_docx_template("' + templateName + '") }}'; 
     *         } else {
       var textBefore = '{{p include_docx_template("' + templateName + '", ' + templateOptions + ') }}'; 
     *         }

     *         // Read the range text
     *         range.load('text');
     *         
     *         range.insertText(textBefore,'Replace');
     *         
     *         await context.sync();
     *         console.log(`The selected text was ${range.text}.`);
       });
     * } */
    
}

/////////////////////////////////////////////////////////////////////
// Helper functions

// File handling
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, {  }, 
                                         function (result: any) {
                                             if (result.status == "succeeded") {
                                                 // If the getFileAsync call succeeded, then
                                                 // result.value will return a valid File Object.
                                                     var myFile = result.value;
                                                 var sliceCount = myFile.sliceCount;
                                                 var slicesReceived = 0, gotAllSlices = true, docdataSlices : any = [];
                                                 //app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

                                                 // Get the file slices.
                                                     getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
                                             }
                                             else {
                                                 //app.showNotification("Error:", result.error.message);
                                             }
                                         });
}

function getSliceAsync(file: any, nextSlice: any, sliceCount: any, gotAllSlices: any, docdataSlices: any, slicesReceived: any) {
    file.getSliceAsync(nextSlice, function (sliceResult: any) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
                // All slices have been received.
                    file.closeAsync();
                onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
        else {
            gotAllSlices = false;
            file.closeAsync();
            //app.showNotification("getSliceAsync Error:", sliceResult.error.message);
        }
    });
}

function onGotAllSlices(docdataSlices: any) {
    var docdata: any = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

function validateUrl(value: string) {
    return /^https?:\/\/\S/i.test(value);
}

//Office.initialize = () => {
    ReactDOM.render(
        <AddinApp />,
        document.getElementById("app")
    );
//}

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Hello } from "./components/Hello";

import {
  ComboBox,
  IComboBoxOption,
  VirtualizedComboBox
} from 'office-ui-fabric-react/lib/ComboBox';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IComboBox } from 'office-ui-fabric-react/lib/components/ComboBox/ComboBox.types';
import {
  assign
} from 'office-ui-fabric-react/lib/Utilities';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

initializeIcons();

export class ComboBoxBasicExample extends React.Component<{}, {
  // For controled single select
  options: IComboBoxOption[];
  selectedOptionKey?: string | number;
  value?: string;

  // For controled multi select
  optionsMulti: IComboBoxOption[];
  selectedOptionKeys?: string[];
  valueMulti?: string;
}> {
  private _testOptions =
	[
	    { key: 'A', text: 'Arial Black' },
	    { key: 'B', text: 'Times New Roman' },
	    { key: 'C', text: 'Comic Sans MS' },
	];

  private scaleOptions: IComboBoxOption[] = [];
  private _basicCombobox: IComboBox;

  constructor(props: {}) {
    super(props);
    this.state = {
      options: [],
      optionsMulti: [],
      selectedOptionKey: undefined,
      value: 'Times New Roman',
      valueMulti: 'Times New Roman'
    };
  }

  public render(): JSX.Element {
    const { options, selectedOptionKey, value } = this.state;
    const { optionsMulti } = this.state;

    return (
      <div className='ms-ComboBoxBasicExample'>

        <ComboBox
          defaultSelectedKey='C'
          label='Basic uncontrolled example (allowFreeform: T, AutoComplete: T):'
          id='Basicdrop1'
          ariaLabel='Basic ComboBox example'
          allowFreeform={ true }
          autoComplete='on'
          options={ this._testOptions }
          onRenderOption={ this._onRenderOption }
          componentRef={ this._basicComboBoxComponentRef }
          // tslint:disable:jsx-no-lambda
          onFocus={ () => console.log('onFocus called') }
          onBlur={ () => console.log('onBlur called') }
          onMenuOpen={ () => console.log('ComboBox menu opened') }
          onPendingValueChanged={ (option, pendingIndex, pendingValue) => console.log('Preview value was changed. Pending index: ' + pendingIndex + '. Pending value: ' + pendingValue) }
        // tslint:enable:jsx-no-lambda
        />

        <DefaultButton
          text='Set focus'
          onClick={ this._basicComboBoxOnClick }
        />
      </div>
    );
  }

  private _onRenderOption = (item: IComboBoxOption): JSX.Element => {
    return <span className={ 'ms-ComboBox-optionText' }>{ item.text }</span>;
  }

  private _getOptions = (currentOptions: IComboBoxOption[]): IComboBoxOption[] => {
    return this.state.options;
  }

  private _onChanged = (option: IComboBoxOption, index: number, value: string): void => {
    console.log('_onChanged() is called: option = ' + JSON.stringify(option));
    if (option !== undefined) {
      this.setState({
        selectedOptionKey: option.key,
        value: undefined
      });
    } else if (index !== undefined && index >= 0 && index < this.state.options.length) {
      this.setState({
        selectedOptionKey: this.state.options[index].key,
        value: undefined
      });
    } else if (value !== undefined) {
      const newOption: IComboBoxOption = { key: value, text: value };

      this.setState({
        options: [...this.state.options, newOption],
        selectedOptionKey: newOption.key,
        value: undefined
      });
    }
  }

  private _updateSelectedOptionKeys = (selectedKeys: string[], option: IComboBoxOption): string[] => {
    if (selectedKeys && option) {
      const index = selectedKeys.indexOf(option.key as string);
      if (option.selected && index < 0) {
        selectedKeys.push(option.key as string);
      } else {
        selectedKeys.splice(index, 1);
      }
    }
    return selectedKeys;
  }

  private _basicComboBoxOnClick = (): void => {
    this._basicCombobox.focus(true);
  }

  private _basicComboBoxComponentRef = (component: IComboBox): void => {
    this._basicCombobox = component;
  }
}

ReactDOM.render(
    <ComboBoxBasicExample />,
    document.getElementById("example")
);

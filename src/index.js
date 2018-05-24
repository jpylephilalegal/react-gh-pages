import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
//import App from './App';
import registerServiceWorker from './registerServiceWorker';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { initializeIcons } from '@uifabric/icons';

initializeIcons();

const MyPage = () => (
  <Fabric>
    <DefaultButton>
    I am a button.
    <img src="static/media/cat-in-circle-32.png" />
    </DefaultButton>
  </Fabric>
);

ReactDOM.render(<MyPage />, document.getElementById('root'));
registerServiceWorker();

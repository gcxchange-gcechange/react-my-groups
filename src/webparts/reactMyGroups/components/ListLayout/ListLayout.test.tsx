import * as React from 'react';
import { configure, mount, ReactWrapper } from 'enzyme';
import {  ISize } from 'office-ui-fabric-react/lib/Utilities';
import * as Adapter from 'enzyme-adapter-react-16';

configure({ adapter: new Adapter() });

import { ListLayout } from './ListLayout';

describe('List Layout basic render test', () => {

  it('should render even if empty', () => {
    var reactComponent = mount(React.createElement(
        ListLayout,
        {
          items: [],
          onRenderListItem: (item: any, finalSize: ISize, isCompact: boolean) => (<div></div>)
        }
    ));
    
    const element = reactComponent.find("FocusZone");
    expect(element.length).toBeGreaterThan(0);

    reactComponent.unmount();
  });
});


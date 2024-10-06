import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export interface GetDataPageBodyProps {
  getFileNames: () => {};
  logout: () => {};
  createTestMailFolder: () => {};
  createFamilyItem: () => {}; // Add this line
  deleteFamilyItem: () => {}; // Pa152
  queryContainer: () => {}; // Pc7bc
}

export default class GetDataPageBody extends React.Component<GetDataPageBodyProps> {
  render() {
    const { getFileNames, logout, createTestMailFolder, createFamilyItem, deleteFamilyItem, queryContainer } = this.props;

    return (
      <div className='ms-welcome__main'>
        <h2 className='ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20'>
          You can add data from Microsoft Graph to the reply of an email on this page.
        </h2>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={getFileNames}
        >
          Get File Names
        </Button>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={createTestMailFolder}
        >
          Create Test Mail Folder
        </Button>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={createFamilyItem} // Update this line
        >
          Create in DB
        </Button>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={deleteFamilyItem} // P50e8
        >
          Remove from DB
        </Button>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={queryContainer} // Pc7bc
        >
          QueryContainer from DB
        </Button>
        <Button
          className='ms-welcome__action'
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={logout}
        >
          Sign out from Office 365
        </Button>
      </div>
    );
  }
}

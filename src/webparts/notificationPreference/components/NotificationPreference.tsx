import * as React from 'react';
import styles from './NotificationPreference.module.scss';
import type { INotificationPreferenceProps, INotificationPreferenceState } from '../interfaces';
import { ChoiceGroup, DirectionalHint, IButtonProps, IChoiceGroupOption, IChoiceGroupOptionStyles, IconButton, IIconProps, Label, MessageBar, ProgressIndicator, Spinner, SpinnerSize, TeachingBubble } from '@fluentui/react';
import { ClapSpinner, PushSpinner } from 'react-spinners-kit';
import { DMSService } from '../services';

export default class NotificationPreference extends React.Component<INotificationPreferenceProps, INotificationPreferenceState> {

  private _Service: DMSService;
  constructor(props: INotificationPreferenceProps) {
    super(props);
    this.state = {
      notificationPreferenceKey: "",
      notificationPreferenceValue: "",
      currentUserId: 0,
      currentUserLoginName: "",
      defaultPreference: this.props.defaultPreferenceText,
      currentPreferenceItemID: 0,
      showMessage: "none",
      message: "",
      messageMode: 4 //Success mode - 4
    };
    this._Service = new DMSService(this.props.context);
    this._onClose = this._onClose.bind(this);
    this.selectPreference = this.selectPreference.bind(this);

  }

  private selectPreference = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: IChoiceGroupOption) => {
    // Updating current user's preference.
    try {
      // Adding as new entry if there is no preference set for the user
      if (this.state.currentPreferenceItemID == 0) {
        const notitem = {
          Title: this.state.currentUserLoginName,
          EmailUserId: this.state.currentUserId,
          Preference: option.key
        }
        this._Service.createNewItem(this.props.hubSiteUrl, this.props.notificationPrefListName, notitem)
          /* sp.web.getList("/sites/" + this.props.hubSiteUrl + "/Lists/" + this.props.notificationPrefListName).items.add
            ({
              Title: this.state.currentUserLoginName,
              EmailUserId: this.state.currentUserId,
              Preference: option.key
            }) */
          .then(currentItemID => {
            this.setState({
              defaultPreference: option.key,
              currentPreferenceItemID: currentItemID.data.ID
            });
          });
      }
      //Upadting preference with the selected value
      else {
        const prefitem = {
          Preference: option.key
        }
        this._Service.updateItemById(this.props.hubSiteUrl, this.props.notificationPrefListName, this.state.currentPreferenceItemID, prefitem)
        /* sp.web.getList("/sites/" + this.props.hubSiteUrl + "/Lists/" + this.props.notificationPrefListName).items.getById(this.state.currentPreferenceItemID).update
          ({
            Preference: option.key
          }); */
      }
      //Reading User messages for 'NotificationPreference' page.
      this._Service.getSelectOrderByFilter(this.props.hubSiteUrl, this.props.userMessageSettings, "Title,Message", "ID", "PageName eq 'NotificationPreference'")
      // this._Service.getItemsFromUserMsgSettingsNP(this.props.hubSiteUrl, this.props.userMessageSettings)
        /* sp.web.getList("/sites/" + this.props.hubSiteUrl + "/Lists/" + this.props.userMessageSettings).items.select("Title,Message").orderBy("ID")
          .filter("PageName eq 'NotificationPreference'").get() */
        .then(userMessages => {
          if (userMessages.length > 0) {
            if (option.key == this.props.noEmail)
              this.setState({
                message: userMessages[0]['Message'],
                defaultPreference: option.key,
                showMessage: ""
              });
            else if (option.key == this.props.sendForCriticalDocuments)
              this.setState({
                defaultPreference: option.key,
                message: userMessages[1]['Message'],
                showMessage: ""
              });
            else if (option.key == this.props.sendForAllDocuments)
              this.setState({
                defaultPreference: option.key,
                message: userMessages[2]['Message'],
                showMessage: ""
              });
          }
        });
    }
    catch (err) {
      this.setState({
        messageMode: 1,// Error mode - 1
        message: err
      });
    }
  }

  public async componentDidMount() {
    //Getting current user's email
    let currentUser = await this._Service.getCurrentUser()
    //let currentUser = await sp.web.currentUser();
    await this.GetCurrentUserDetails();
    // Getting current uuser's preference if already set.
    const notificationPreference: any[] = await this._Service.getSelectExpandFilter(this.props.hubSiteUrl, this.props.notificationPrefListName, "ID,Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail", "EmailUser", "EmailUser/EMail eq '" + currentUser.Email + "'")
    // const notificationPreference: any[] = await this._Service.getNotificationPref(this.props.hubSiteUrl, this.props.notificationPrefListName, currentUser.Email)
    //const notificationPreference: any[] = await sp.web.getList("/sites/" + this.props.hubSiteUrl + "/Lists/" + this.props.notificationPrefListName).items.select("ID,Preference,EmailUser/ID,EmailUser/Title,EmailUser/EMail").expand("EmailUser").filter("EmailUser/EMail eq '" + currentUser.Email + "'").get();
    if (notificationPreference.length > 0) {
      this.setState({
        defaultPreference: notificationPreference[0].Preference,
        currentPreferenceItemID: notificationPreference[0].ID
      });
    }
    console.log(this.state.defaultPreference);
  }
  protected async GetCurrentUserDetails() {
    let currentUser = await this._Service.getCurrentUser()
    //let currentUser = await sp.web.currentUser();
    this.setState({
      currentUserId: currentUser.Id,
      currentUserLoginName: currentUser.Title,
    });
  }
  private _onClose = () => {
    // alert(window.location.protocol +"//"+window.location.hostname+"/sites/"+this.props.hubSiteUrl);
    window.location.replace(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubSiteUrl);
  }

  public render(): React.ReactElement<INotificationPreferenceProps> {
    // Icon setting for close button
    const cancelIcon: IIconProps = { iconName: 'Cancel' };
    // Setting options for checkboxgroup
    const options: IChoiceGroupOption[] = [
      { key: this.props.noEmail, text: this.props.noEmailText, iconProps: { iconName: 'MailUndelivered' }, title: this.props.noEmailText },
      { key: this.props.sendForCriticalDocuments, text: this.props.sendForCriticalDocumentsText, iconProps: { iconName: 'MailAlert' }, title: this.props.sendForCriticalDocumentsText },
      { key: this.props.sendForAllDocuments, text: this.props.sendForAllDocumentsText, iconProps: { iconName: 'MailCheck' }, title: this.props.sendForAllDocumentsText },
    ];
    const searchBoxStyles: Partial<IChoiceGroupOptionStyles> = {
      root: { innerWidth: "20px", innerHeight: "20px" }
    };
    // Setting style for checkbox
    const preferenceCheckBoxStyle: Partial<IChoiceGroupOptionStyles> = {
      root: { innerWidth: "120px", innerHeight: "40px" }
    };

    return (
      <section className={`${styles.notificationPreference}`}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ textAlign: 'right' }}><IconButton iconProps={cancelIcon} title="Close" ariaLabel=" " onClick={this._onClose} style={{ marginTop: "0px" }} /></div>
          <div style={{ fontSize: '24px', fontWeight: "lighter" }}> {this.props.description} </div>  <br></br>
          <div>Your current preference : {this.state.defaultPreference}</div>
          <div style={{ marginTop: "20px" }} className={styles['labelWrapper-124']}>
            <ChoiceGroup defaultSelectedKey={this.state.defaultPreference} defaultValue={this.state.defaultPreference} options={options} style={{ marginLeft: "252px", marginTop: "12px" }} styles={preferenceCheckBoxStyle} id={'targetChoice'} onChange={this.selectPreference} />
          </div>
          <div style={{ display: this.state.showMessage }}>
            <MessageBar messageBarType={this.state.messageMode} isMultiline={false}>{this.state.message}</MessageBar>
          </div>
        </div>
      </section>
    );
  }
}

import * as React from 'react';

import Loader from 'react-loader-spinner';

import styles from './OneDriveForm.module.scss';
import { IOneDriveFormProps } from './IOneDriveFormProps';

import { getCurrentUserLookupId, getListItemsByUserId, createOneDriveMigrationRequest, getListItemEntityTypeName, getItemsFromCompletedListByUser } from '../services/sp-rest';

const initialState = {
  userId: null,
  alreadySubmitted: false,
  alreadyCompleted: false,
  finishedSubmitting: false,
  buttonDisabled: true,
  isLoading: true,
  isLoadingTypeName: true,
  typeName: '',
  didError: false
};

type State = Readonly<typeof initialState>;

export default class OneDriveForm extends React.Component<IOneDriveFormProps, {}> {
  public readonly state: State = initialState;
  private listItemEntityTypeName: string;


  private checkIfSubmitted() {
    getListItemsByUserId(this.props.webpartContext, this.state.userId).then((response) => {

      response.json().then(responseJSON => {
        if (responseJSON.value.length > 0) {
          // If items return, then the user has already submitted.
          this.setState({
            isLoading: false,
            alreadySubmitted: true
          });
        } else {
          let usersName = this.props.webpartContext.pageContext.user.loginName.replace('i:0#.f|membership|', '').split("@")[0];

          getItemsFromCompletedListByUser(this.props.webpartContext, usersName).then((response2) => {
            response2.json().then(responseJSON2 => {
              if (responseJSON2.value.length > 0) {
                // If items return, then the user has already submitted.
                this.setState({
                  isLoading: false,
                  alreadyCompleted: true
                });
              } else {
                // If no items return, the user has not submitted yet.
                this.setState({
                  isLoading: false,
                  buttonDisabled: false
                });
              }
            });
          });

        }
      }).catch(error => {
        // Not sure why this would ever error, but handling it just in case.
        this.setState({
          didError: true
        });
      });
    }).catch(error => {
      // Handle failed calls. E.G. The list cannot be found or the user does not have access to the list.
      this.setState({
        didError: true
      });
    });
  }

  // Handle some setup after the component mounts
  public componentDidMount() {

    // Get the current user lookup ID
    getCurrentUserLookupId(this.props.webpartContext.pageContext.user.loginName, this.props.webpartContext.pageContext.site.absoluteUrl, this.props.webpartContext.spHttpClient)
      .then(result => {
        result.json().then(rJson => {
          this.setState({
            userId: rJson.Id
          });

          // after setting the user's ID, check if they have already submitted the form.
          this.checkIfSubmitted();
        });
      })
      .catch(error => {
        // handle if we cannot get the user's ID.
        this.setState({
          didError: true
        });
      });

    getListItemEntityTypeName(this.props.webpartContext.pageContext.site.absoluteUrl, this.props.webpartContext.spHttpClient).then(response => {

      if (response.ok) {
        response.json().then(rJson => {
          this.setState({
            isLoadingTypeName: false,
            typeName: rJson.ListItemEntityTypeFullName
          });
        });
      } else {
        this.setState({
          didError: true
        });
      }
    });
  }


  private handleRequestButtonClick() {
    // Set the state as loading before trying to create the request.
    this.setState({
      isLoading: true
    });

    // Try to create the request.
    createOneDriveMigrationRequest(this.props.webpartContext, this.state.userId, this.state.typeName).then((result) => {
      this.setState({
        isLoading: false,
        finishedSubmitting: true
      });
    }).catch(err => {
      this.setState({
        isLoading: false,
        didError: true
      });
    });
  }

  public render(): React.ReactElement<IOneDriveFormProps> {
    let description = this.props.description;

    window['description'] = description;

    if (!description || (description.trim && !description.trim() && description.trim().length === 0)) {

      description = "Click the button if you are ready to have your OneDrive content moved. This will create a request that will be process in the near future.";
    } else {

    }

    if (this.state.didError) {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}>An error occurred. Please contact the WCM Team.</p>
              </div>
            </div>
          </div>
        </div>
      );
    } else if (this.state.isLoading || this.state.isLoadingTypeName) {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}></p>
              </div>
              <div className={styles.loader}>
                <Loader className={styles.loader}
                  type="Oval"
                  color="#00BFFF"
                  height="100"
                  width="100"
                />
              </div>
            </div>
          </div>
        </div>
      );
    } else if (this.state.alreadySubmitted) {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}>You have already submitted the request.</p>
              </div>
            </div>
          </div>
        </div>
      );
    } else if (this.state.alreadyCompleted) {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}>Your request has already been completed.</p>
              </div>
            </div>
          </div>
        </div>
      );
    } else if (this.state.finishedSubmitting) {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}>Your request was successfully submitted.</p>
              </div>
            </div>
          </div>
        </div>
      );
    } else {
      return (
        <div className={styles.oneDriveForm}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>One Drive Migration Form</span>
                {/* <p className={styles.subTitle}>Subtitle goes here.</p> */}
                <p className={styles.description}>{description}</p>
                <button disabled={this.state.buttonDisabled} onClick={() => { this.handleRequestButtonClick(); }} className={styles.button}>
                  <span className={styles.label}>Request OneDrive Migration</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      );
    }
  }
}

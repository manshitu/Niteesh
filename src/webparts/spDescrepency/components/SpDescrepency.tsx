import * as React from 'react';
import styles from './SpDescrepency.module.scss';
import type { ISpDescrepencyProps } from './ISpDescrepencyProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpDescrepency extends React.Component<ISpDescrepencyProps> {
  public render(): React.ReactElement<ISpDescrepencyProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section
        className={`${styles.spDescrepency} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part name: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to Descrepency Management!</h3>
          <p>
            This webpart will be used for descrepency Management, Please upload
            excel sheet for the employee data.
          </p>
          <h4>Browse and select excel file to read and compare data:</h4>
          <div>
            <input type="file" accept=".xlsx" /><br /><br />
            <button>Upload and Compare</button>
          </div>
        </div>
      </section>
    );
  }
}
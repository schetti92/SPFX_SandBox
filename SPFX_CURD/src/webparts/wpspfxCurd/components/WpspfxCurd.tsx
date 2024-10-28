import * as React from "react";
import styles from "./WpspfxCurd.module.scss";
import type { IWpspfxCurdProps } from "./IWpspfxCurdProps";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import {
  PrimaryButton,
  TextField,
  Stack,
  DatePicker,
  IStackTokens,
  ComboBox,
  IComboBox,
  IComboBoxOption,
} from "@fluentui/react";

import { getSP } from "../pnpjsConfig";
import { SPFI } from "@pnp/sp";

// Define stackTokens to specify spacing between Stack children
const stackTokens: IStackTokens = { childrenGap: 10 };
const titleOptions: IComboBoxOption[] = [
  { key: "Mr", text: "Mr" },
  { key: "Miss", text: "Miss" },
  { key: "Mrs", text: "Mrs" },
  { key: "Ms", text: "Ms" },
  { key: "Other", text: "Other" },
];

export interface IWpspfxCurdState extends IWpspfxCurdStateErrorHandle {
  showOtherTitle: boolean;
  showDeleteButton: boolean;
  showUpdateButton: boolean;
  isOtherTitle: boolean;
  firstName: string;
  lastName: string;
  dob: Date | null;
  address: string;
  contactNumber: string;
  email: string;
  title: string;
  otherTitle: string;
  searchId: number;
}

export interface IWpspfxCurdStateErrorHandle {
  firstNameError: string;
  contactNumberError: string;
}

export default class WpspfxCurd extends React.Component<
  IWpspfxCurdProps,
  IWpspfxCurdState,
  IWpspfxCurdStateErrorHandle
> {
  public _sp: SPFI;

  constructor(props: IWpspfxCurdProps) {
    super(props);
    this._sp = getSP();
    this.state = {
      showOtherTitle: false,
      showDeleteButton: false,
      showUpdateButton: false,
      isOtherTitle: false,
      firstName: "",
      lastName: "",
      dob: null,
      address: "",
      contactNumber: "",
      email: "",
      title: "",
      otherTitle: "",
      searchId: 0,
      firstNameError: "",
      contactNumberError: "",
    };
  }

  // Handle title change
  private onTitleChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption
  ): void => {
    this.setState({
      title: option?.key.toString() || "",
      isOtherTitle: option?.key === "Other",
      showOtherTitle: option?.key === "Other",
    });
  };

  // Handle input changes
  private handleInputChange =
    (field: keyof IWpspfxCurdState) =>
    (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
      this.setState({ [field]: event.currentTarget.value } as any);
    };

  // Handle date change
  private onDateChange = (date: Date | null): void => {
    this.setState({ dob: date });
  };

  private saveData = async (): Promise<void> => {
    const { listName } = this.props;
    const {
      firstName,
      lastName,
      dob,
      address,
      contactNumber,
      email,
      title,
      otherTitle,
    } = this.state;

    let hasError = false;

    this.setState({ firstNameError: "", contactNumberError: "" });

    // Validation
    if (!firstName) {
      this.setState({ firstNameError: "First Name is required." });
      hasError = true;
    }

    if (!contactNumber) {
      this.setState({ contactNumberError: "Contact Number is required." });
      hasError = true;
    }

    if (hasError) {
      return; // Stop execution if there are errors
    }

    try {
      console.log("Title:", title, "Other Title:", otherTitle); // Log both values
      await this._sp.web.lists.getByTitle(listName).items.add({
        Title: title === "Other" ? otherTitle : title, // Use the other title if selected
        FirstName: firstName,
        LastName: lastName,
        DOB: dob?.toISOString(), // Convert date to ISO format
        Address: address,
        ContactNumber: contactNumber,
        Email: email,
        OtherTitle: title === "Other",
      });
      alert("Data saved successfully!");
      this.onReset();
    } catch (error) {
      console.error("Error saving data:", error);
      alert("Failed to save data. Please try again.");
    }
  };

  // Handle search by ListItem ID
  private onSearch = async (id: number): Promise<void> => {
    if (!id) {
      alert("Please enter a valid ID.");
      this.onReset();
      return;
    }
    this.onReset();

    try {
      console.error("Searching for ID:", id);

      const item: any = await this._sp.web.lists
        .getByTitle(this.props.listName)
        .items.getById(id)();

      console.log("Retrieved item:", item?.FirstName);

      if (item) {
        let extTitle: string;
        let otherTitle: string;

        console.log(item.OtherTitle);

        if (item.OtherTitle === true) {
          extTitle = "Other";
          otherTitle = item.Title; // Assign "Other" to otherTitle
          this.setState({ showOtherTitle: true });
        } else {
          this.setState({ showOtherTitle: false });
          extTitle = item.Title; // Use the existing title
          otherTitle = ""; // Set otherTitle to an empty string
        }

        this.setState({
          showDeleteButton: true,
          showUpdateButton: true,
          firstName: item.FirstName || "",
          lastName: item.LastName || "",
          dob: item.DOB ? new Date(item.DOB) : null,
          address: item.Address || "",
          contactNumber: item.ContactNumber || "",
          email: item.Email || "",
          title: extTitle, // Set the computed title
          otherTitle: otherTitle, // Set the computed otherTitle
          searchId: id,
        });
      }

      console.log("Title :", this.state.title);
      console.log("Other Title :", this.state.otherTitle);
    } catch (error) {
      console.error("Error fetching data:", error);

      // Show alert if the error is due to an invalid ID
      alert("No matching records found.");

      // Clear fields if no records found
      this.onReset();
    }
  };

  private onReset = async (): Promise<void> => {
    try {
      this.setState({
        showDeleteButton: false,
        showUpdateButton: false,
        firstName: "",
        lastName: "",
        dob: null,
        address: "",
        contactNumber: "",
        email: "",
        title: "",
        otherTitle: "",
        searchId: 0,
        showOtherTitle: false,
        firstNameError: "",
        contactNumberError: "",
      });
    } catch (error) {
      console.error("Error clear data:", error);
    }
  };

  private deleteData = async (id: number): Promise<void> => {
    if (!id) {
      console.error("Please enter a valid ID.");
      return;
    }
    const { listName } = this.props;
    try {
      const list = this._sp.web.lists.getByTitle(listName);
      await list.items.getById(id).delete();
      alert("Data deleted successfully!");
    } catch (error) {
      console.error("Error deleteing data:", error);
      alert("Failed to delete data. Please try again.");
    }
  };

  private updateData = async (id: number): Promise<void> => {
    if (!id) {
      console.error("Please enter a valid ID.");
      return;
    }
    const { listName } = this.props;
    try {
      const list = this._sp.web.lists.getByTitle(listName);
      const {
        firstName,
        lastName,
        dob,
        address,
        contactNumber,
        email,
        title,
        otherTitle,
      } = this.state;

      let hasError = false;

      this.setState({ firstNameError: "", contactNumberError: "" });

      // Validation
      if (!firstName) {
        this.setState({ firstNameError: "First Name is required." });
        hasError = true;
      }

      if (!contactNumber) {
        this.setState({ contactNumberError: "Contact Number is required." });
        hasError = true;
      }

      if (hasError) {
        return; // Stop execution if there are errors
      }

      await list.items.getById(id).update({
        Title: title === "Other" ? otherTitle : title, // Use the other title if selected
        FirstName: firstName,
        LastName: lastName,
        DOB: dob?.toISOString(), // Convert date to ISO format
        Address: address,
        ContactNumber: contactNumber,
        Email: email,
        OtherTitle: title === "Other",
      });
      alert("Item updated successfully!");
      this.onReset();
    } catch (error) {
      console.error("Error updated data:", error);
      alert("Failed to updated data. Please try again.");
    }
  };

  public render(): React.ReactElement<IWpspfxCurdProps> {
    return (
      <section className={`${styles.wpspfxCurd}`}>
        <div>
          <h1>Student Registration Details</h1>
          <h3>List Name: {this.props.listName}</h3>
        </div>
        {this.props.listName && (
          <div>
            <Stack tokens={stackTokens} styles={{ root: { width: "100%" } }}>
              <Stack horizontal tokens={stackTokens}>
                <SearchBox
                  placeholder="Search by Student Registration #"
                  onSearch={async (newValue) => {
                    const id = parseInt(newValue, 10); // Convert input to a number
                    await this.onSearch(id); // Call the onSearch method
                  }}
                  styles={{ root: { flexGrow: 1 } }}
                />
                {this.state.showDeleteButton && (
                  <PrimaryButton onClick={this.onReset} text="Reset" />
                )}
              </Stack>
              <Stack horizontal tokens={stackTokens}>
                <ComboBox
                  label="Title"
                  options={titleOptions}
                  onChange={this.onTitleChange}
                  selectedKey={this.state.title}
                  styles={{ root: { flexGrow: 1 } }}
                />
                {this.state.showOtherTitle && (
                  <TextField
                    label="Please specify"
                    onChange={this.handleInputChange("otherTitle")} // Update other title state
                    value={this.state.otherTitle}
                    styles={{ root: { flexGrow: 1 } }}
                  />
                )}
              </Stack>
              <Stack horizontal tokens={stackTokens}>
                <TextField
                  label="First Name"
                  required
                  errorMessage={this.state.firstNameError}
                  onChange={this.handleInputChange("firstName")} // Update first name state
                  value={this.state.firstName} // Bind value to state
                  styles={{ root: { flexGrow: 1 } }}
                />
                <TextField
                  label="Last Name"
                  onChange={this.handleInputChange("lastName")} // Update last name state
                  value={this.state.lastName} // B
                  styles={{ root: { flexGrow: 1 } }}
                />
              </Stack>
              <Stack>
                <DatePicker
                  label="Date Of Birth"
                  value={this.state.dob ?? undefined}
                  onSelectDate={this.onDateChange} // Handle date selection
                  styles={{ root: { width: "50%" } }}
                />
              </Stack>
              <Stack>
                <TextField
                  label="Address"
                  multiline
                  rows={3}
                  value={this.state.address} // B
                  onChange={this.handleInputChange("address")} // Update address state
                />
              </Stack>
              <Stack horizontal tokens={stackTokens}>
                <TextField
                  label="Contact Number"
                  required
                  errorMessage={this.state.contactNumberError}
                  onChange={this.handleInputChange("contactNumber")} // Update contact number state
                  value={this.state.contactNumber} // B
                  styles={{ root: { flexGrow: 1 } }}
                />
                <TextField
                  label="Email"
                  onChange={this.handleInputChange("email")} // Update email state
                  value={this.state.email} // B
                  styles={{ root: { flexGrow: 1 } }}
                />
              </Stack>
              <Stack horizontal tokens={stackTokens}>
                {!this.state.showUpdateButton && (
                  <PrimaryButton text="Submit" onClick={this.saveData} />
                )}
                {this.state.showUpdateButton && (
                  <PrimaryButton
                    text="Update"
                    onClick={() => this.updateData(this.state.searchId)}
                  />
                )}
                <PrimaryButton text="Close" />
                {this.state.showDeleteButton && (
                  <PrimaryButton
                    text="Delete"
                    onClick={() => this.deleteData(this.state.searchId)}
                  />
                )}
              </Stack>
            </Stack>
          </div>
        )}
      </section>
    );
  }
}

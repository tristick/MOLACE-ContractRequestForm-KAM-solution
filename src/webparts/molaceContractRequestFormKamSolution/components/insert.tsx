import React, { Component } from 'react';
import { TextField, Dropdown, IconButton, TooltipHost } from 'office-ui-fabric-react';
import styles from "./Trpreqfrm.module.scss"

interface DynamicTableProps {
  items: any[];
  onUpdateItems: (items: any[]) => void;
 
}

// Define the state type
interface DynamicTableState {
  newDataRow: {
    pol: string;
    pod: string;
    volume:  number | '';
    baseFreight: string;
    surcharges: string;
  };
  rows: Array<{
    pol: string;
    pod: string;
    volume:  number | '';
    baseFreight: string;
    surcharges: string;
    isEditing: boolean;
  }>;
  editingData: {
    pol: string;
    pod: string;
    volume:  number | '';
    baseFreight: string;
    surcharges: string;
  };
  currentlyEditingIndex: number | null;
  showEditTooltip: boolean
}

class DynamicTable extends Component<DynamicTableProps, DynamicTableState> {
  constructor(props: DynamicTableProps) {
    super(props);

    this.state = {
      newDataRow: {
        pol: '',
        pod: '',
        volume: 0,
        baseFreight: '',
        surcharges: '-Select-',
      },
      rows: [],
      editingData: {
        pol: '',
        pod: '',
        volume: 0,
        baseFreight: '',
        surcharges: '-Select-',
      },
      currentlyEditingIndex: null, // Initialize to null
      showEditTooltip: false,
    };
   
  }

  handleInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, fieldName: keyof DynamicTableState['newDataRow'], dataKey?: string) => {
    const { newDataRow, editingData } = this.state;
    const targetValue = (event.target as HTMLInputElement).value;
    if (fieldName === 'volume') {
      // Validate and parse the volume as a number
      const volume = /^\d+$/.test(targetValue) ? parseInt(targetValue, 10) : '';

      if (dataKey) {
        editingData[fieldName] = volume;
        this.setState({ editingData });
      } else {
        newDataRow[fieldName] = volume;
        this.setState({ newDataRow });
      }
    } else {
      if (dataKey) {
        editingData[fieldName] = targetValue;
        this.setState({ editingData });
      } else {
        newDataRow[fieldName] = targetValue;
        this.setState({ newDataRow });
      }
    }
  };

  handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option: { key: any; }, dataKey?: string) => {
    const { newDataRow, editingData } = this.state;
    const selectedKey = option.key;

    if (dataKey) {
      editingData.surcharges = selectedKey;
      this.setState({ editingData });
    } else {
      newDataRow.surcharges = selectedKey;
      this.setState({ newDataRow });
    }
  };

  addRow = () => {
    const { newDataRow, rows } = this.state;
    if (!newDataRow.pol || !newDataRow.pod || newDataRow.volume === ''  || !newDataRow.baseFreight || newDataRow.surcharges === '-Select-') {
      return;
    }
    const updatedRows = [...rows, { ...newDataRow, isEditing: false }];
    const emptyDataRow = {
      pol: '',
      pod: '',
      volume: 0,
      baseFreight: '',
      surcharges: '-Select-',
    };

    const updatedItems = [...this.props.items, { ...newDataRow }];

    this.setState({
      rows: updatedRows,
      newDataRow: emptyDataRow,
    });

    this.props.onUpdateItems(updatedItems);
    
  };

  editRow = (index: number) => {
    if (this.state.currentlyEditingIndex !== null) {
      // If another row is already being edited, do not allow editing another row
      
      return;
    }

    const { rows } = this.state;
    const editingData = { ...rows[index] };
    
    const updatedRows = [...rows];
    updatedRows[index].isEditing = true;
    this.setState({ rows: updatedRows, editingData, currentlyEditingIndex: index  });
  };

  saveEditedRow = (index: number) => {
    const { rows, editingData } = this.state;

     // Check if any control in the editingData is empty
  if (
    editingData.pol.trim() === '' ||
    editingData.pod.trim() === '' ||
    editingData.volume.toString().trim() === '' ||
    editingData.baseFreight.trim() === '' ||
    editingData.surcharges.trim() === '-Select-'
  ) {
    // Display an error message or take appropriate action (e.g., show a tooltip)
    
    this.setState({showEditTooltip:true});
    console.log(this.state.showEditTooltip)
    return;
  }
    const updatedRows = [...rows];
    updatedRows[index] = { ...editingData, isEditing: false };
    this.setState({ rows: updatedRows, editingData: {
      pol: '',
      pod: '',
      volume: '',
      baseFreight: '',
      surcharges: ''
    }, currentlyEditingIndex: null }, );
     // Update the items state when saving an edited row
     const updatedItems = updatedRows.map((row) => ({ ...row }));
     this.props.onUpdateItems(updatedItems);
     this.setState({ showEditTooltip: false });
    };

/*   cancelEdit = (index: number) => {
    const { rows } = this.state;
    const updatedRows = [...rows];
    updatedRows[index].isEditing = false;
    this.setState({ rows: updatedRows, editingData: {
      pol: '',
      pod: '',
      volume: '',
      baseFreight: '',
      surcharges: ''
    } });
    // Update the items state when canceling the edit
    const updatedItems = updatedRows.map((row) => ({ ...row }));
    this.props.onUpdateItems(updatedItems);
  }; */

  deleteRow = (index: number) => {
    const { rows } = this.state;
    const updatedRows = [...rows];
    updatedRows.splice(index, 1);
    this.setState({ rows: updatedRows });
    
    // Update the items state when deleting a row
    const updatedItems = updatedRows.map((row) => ({ ...row }));
    this.props.onUpdateItems(updatedItems);
    
  
  };

  calculateTotalVolume = () => {
    const { rows } = this.state;
    const newTotalVolume =  rows.reduce((acc, row) => acc + (row.volume || 0), 0);
    
    return newTotalVolume;
  };


  render() {
    const { newDataRow, rows, editingData,currentlyEditingIndex, showEditTooltip  } = this.state;

    return (
      <div>
        <table className="table">
          
          
       <thead>
            <tr>
              <th style={{ width: '180px' }}>POL<span className={styles.required}> *</span></th>
              <th style={{ width: '180px' }}>POD<span className={styles.required}> *</span></th>
              <th style={{ width: '180px' }}>Volume/Year<span className={styles.required}> *</span></th>
              <th style={{ width: '180px' }}>Base Freight<span className={styles.required}> *</span></th>
              <th style={{ width: '180px' }}>Surcharges<span className={styles.required}> *</span></th>
             
            </tr>
          </thead>
        {/* Render your input fields and buttons */}
        <tbody>
          <tr><td >
        <TextField
       
          value={newDataRow.pol}
          onChange={(event) => this.handleInputChange(event, 'pol')}
        /></td><td>
         <TextField
         
          value={newDataRow.pod}
          onChange={(event) => this.handleInputChange(event, 'pod')}
        /></td><td>
         <TextField
          placeholder='Enter Numbers Only'
          value={newDataRow.volume.toString()}
          onChange={(event) => this.handleInputChange(event, 'volume')}
        /></td><td>
         <TextField
         
          value={newDataRow.baseFreight}
          onChange={(event) => this.handleInputChange(event, 'baseFreight')}
        /></td><td>
        <Dropdown
          
          selectedKey={newDataRow.surcharges}
          options={[
            { key: '-Select-', text: '-Select-' },
            { key: 'Surcharge 1', text: 'Surcharge 1' },
            { key: 'Surcharge 2', text: 'Surcharge 2' },
            // Add more options as needed
          ]}
          onChange={(event, option) => this.handleDropdownChange(event, option)}
        /></td><td>
        <IconButton iconProps={{ iconName: 'Add' }} alt="Insert" onClick={this.addRow} style={{ width: '50px' }}><span>Insert</span> </IconButton>
       
        </td></tr></tbody></table>

       
        {/* Render your table or display the added rows */}
        {rows.length === 0 ? (
          <div className={styles.centered}>
          <div className="no-items-message">No items added.</div>
        </div>
        ) : (
          <table className="table">
          
          <tbody>
            {rows.map((row, index) => (
              <tr key={index}>
                <td style={{ width: '170px' }}>
                  {row.isEditing ? (
                    <TextField
                      value={editingData.pol}
                      onChange={(event) => this.handleInputChange(event, 'pol', 'editingData')}
                    />
                  ) : (
                    <div style={{width: '170px', border: '1px solid #ccc', padding: '4px', whiteSpace: 'nowrap',  // Prevent text from wrapping
                    overflow: 'hidden',    // Hide any overflow
                    textOverflow: 'ellipsis'  // Show ellipsis (...) for overflow
                    }}>{row.pol}</div>
                  )}
                </td>
                {/* Render other fields similarly */}
                  <td style={{ width: '170px' }}>
                  {row.isEditing ? (
                    <TextField value={editingData.pod} onChange={(event) => this.handleInputChange(event, 'pod', 'editingData')} />
                  ) : (
                    <div style={{ width: '170px', border: '1px solid #ccc', padding: '4px' , whiteSpace: 'nowrap',  // Prevent text from wrapping
                    overflow: 'hidden',    // Hide any overflow
                    textOverflow: 'ellipsis'  // Show ellipsis (...) for overflow
                   }}>{row.pod}</div>
                  )}
                </td>
                <td style={{ width: '170px' }}>
                  {row.isEditing ? (
                    <TextField value={editingData.volume.toString()} onChange={(event) => this.handleInputChange(event, 'volume', 'editingData')} />
                  ) : (
                    <div style={{ width: '170px', border: '1px solid #ccc', padding: '4px', whiteSpace: 'nowrap',  // Prevent text from wrapping
                    overflow: 'hidden',    // Hide any overflow
                    textOverflow: 'ellipsis'  // Show ellipsis (...) for overflow
                    }}>{row.volume}</div>
                  )}
                </td>
                <td style={{ width: '170px' }}>
                  {row.isEditing ? (
                    <TextField value={editingData.baseFreight} onChange={(event) => this.handleInputChange(event, 'baseFreight', 'editingData')} />
                  ) : (
                    <div style={{ width: '170px', border: '1px solid #ccc', padding: '4px', whiteSpace: 'nowrap',  // Prevent text from wrapping
                    overflow: 'hidden',    // Hide any overflow
                    textOverflow: 'ellipsis'  // Show ellipsis (...) for overflow
                   }}>{row.baseFreight}</div>
                  )}
                </td>
                <td style={{ width: '170px' }}>
                  {row.isEditing ? (
                    <Dropdown
                      selectedKey={editingData.surcharges}
                      options={[
                        { key: '-Select-', text: '-Select-' },
                        { key: 'Surcharge 1', text: 'Surcharge 1' },
                        { key: 'Surcharge 2', text: 'Surcharge 2' },
                        // Add more options as needed
                      ]}
                      onChange={(event, option) => this.handleDropdownChange(event, option, 'editingData')}
                    />
                  ) : (
                    <div style={{ width: '170px', border: '1px solid #ccc', padding: '4px' }}>{row.surcharges}</div>
                  )}
                </td>
                <td>
                  {row.isEditing ? (
                    <>
                      <IconButton iconProps={{ iconName: 'CheckMark' }} onClick={() => this.saveEditedRow(index)} /><br />
                      {showEditTooltip && (
                    <TooltipHost id="error-tooltip">All fields are mandatory.</TooltipHost>
        )}
                    {/* <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => this.cancelEdit(index)} /> */}
                    </>
                  ) : (
                    <>
                      {/* <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this.editRow(index)} /> */}
                       {/* Use the Tooltip component here */}
                       {currentlyEditingIndex !== null ? (
                            // Display a disabled button with the tooltip for editing
                            <div data-tip
                            data-for="edit-tooltip"
                            //nClick={() => this.setState({ showEditTooltip: true })}
                          >
                              <IconButton disabled iconProps={{ iconName: 'Edit' }} />
                            </div>
                            

                          ) : (
                            <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this.editRow(index)} />
                          )}
                          {/* End of Tooltip component */}
                      <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this.deleteRow(index)} />
                    </>
                  )}
                </td>
              </tr>
            ))}
            <tr><td></td><td></td><td> <div>
          <div><b>Total:<span id = "totalvolumeID">{this.calculateTotalVolume()}</span></b></div>
          
        </div></td></tr>
          </tbody>
        </table>
        )}
        
      </div>
    );
  }
}



export default DynamicTable;

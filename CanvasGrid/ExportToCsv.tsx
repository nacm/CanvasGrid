import * as React from 'react';
import { IconButton, Stack } from 'office-ui-fabric-react';
import { IColumn, IDetailsListProps } from 'office-ui-fabric-react/lib/DetailsList';
import { initializeIcons } from '@uifabric/icons';

export interface ExportToCSVProps extends IDetailsListProps {
}

export class ExportToCSV extends React.Component<ExportToCSVProps> {
  exportToCSV = () => {
    const { items, columns } = this.props;
    const csvContent = `data:text/csv;charset=utf-8,${this.formatCSVData(items, columns)}`;
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', 'data.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  formatCSVData = (items: any[], columns: IColumn[]) => {
    let csvData = '';

    // Write header row
    const headers = columns.map(column => column.name);
    csvData += headers.join(',') + '\n';

    // Write data rows
    items.forEach(item => {
      const row = columns.map(column => {
        const fieldValue = item[column.fieldName];
        return fieldValue !== undefined ? fieldValue : '';
      }).join(',');
      csvData += row + '\n';
    });

    return csvData;
  };

  componentDidMount() {
    initializeIcons();
  }

  render() {
    const { items, columns } = this.props;

    return (
      <Stack horizontal>
        <IconButton
          iconProps={{ iconName: 'ExcelDocument' }}
          title="Export to CSV"
          ariaLabel="Export to CSV"
          onClick={this.exportToCSV}
        />
        <div>
          {/* Render your Fluent UI DetailList component */}
          {/* Pass the 'items' and 'columns' props to the DetailList */}
        </div>
      </Stack>
    );
  }
}

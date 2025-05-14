import * as React from 'react';
import styles from './Firstapp.module.scss';
import type { IFirstappProps } from './IFirstappProps';
import { IListItem } from './interfaces';
import { getSP } from '../pnpjsConfig';
import {
  Label,
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  IconButton,
} from '@fluentui/react';

interface State {
  items: IListItem[];
  loading: boolean;
  title: string;
  body: string;
  letter: string;
  editItemId: number | null;
  currentPage: number;
  itemsPerPage: number;
}

export default class Firstapp extends React.Component<IFirstappProps, State> {
  private _sp = getSP();
  private LIST_NAME = 'FAQ';

  constructor(props: IFirstappProps) {
    super(props);
    this.state = {
      items: [],
      loading: false,
      title: '',
      body: '',
      letter: '',
      editItemId: null,
      currentPage: 1,
      itemsPerPage: 5,
    };
  }

  public componentDidMount(): void {
    void this._loadListItems();
  }

  private _loadListItems = async (): Promise<void> => {
    this.setState({ loading: true });
    try {
      const listItems = await this._sp.web.lists
        .getByTitle(this.LIST_NAME)
        .items.select('Id', 'Title', 'Body', 'Letter')();
      this.setState({ items: listItems, loading: false });
    } catch (error) {
      console.error('Error loading list items:', error);
      this.setState({ loading: false });
    }
  };

  private _handleInputChange = (_: any, newValue?: string, field?: keyof Pick<State, 'title' | 'body' | 'letter'>) => {
    if (field) this.setState({ [field]: newValue || '' } as unknown as Pick<State, keyof State>);
  };

  private _createOrUpdateItem = async () => {
    const { title, body, letter, editItemId } = this.state;

    if (!title) return alert('Title is required');

    try {
      if (editItemId) {
        await this._sp.web.lists.getByTitle(this.LIST_NAME).items.getById(editItemId).update({
          Title: title,
          Body: body,
          Letter: letter,
        });
      } else {
        await this._sp.web.lists.getByTitle(this.LIST_NAME).items.add({
          Title: title,
          Body: body,
          Letter: letter,
        });
      }

      this.setState({ title: '', body: '', letter: '', editItemId: null });
      await this._loadListItems();
    } catch (error) {
      console.error('Error saving item:', error);
    }
  };

  private _editItem = (item: IListItem) => {
    this.setState({
      editItemId: item.Id,
      title: item.Title,
      body: item.Body,
      letter: item.Letter,
    });
  };

  private _deleteItem = async (id: number) => {
    if (confirm('Are you sure you want to delete this item?')) {
      try {
        await this._sp.web.lists.getByTitle(this.LIST_NAME).items.getById(id).delete();
        await this._loadListItems();
      } catch (error) {
        console.error('Error deleting item:', error);
      }
    }
  };

  private _handlePageChange = (pageNumber: number) => {
    this.setState({ currentPage: pageNumber });
  };

  public render(): React.ReactElement<IFirstappProps> {
    const {
      items,
      loading,
      title,
      body,
      letter,
      editItemId,
      currentPage,
      itemsPerPage,
    } = this.state;

    const startIndex = (currentPage - 1) * itemsPerPage;
    const pagedItems = items.slice(startIndex, startIndex + itemsPerPage);
    const totalPages = Math.ceil(items.length / itemsPerPage);

    return (
      <section className={styles.firstapp}>
        <Stack tokens={{ childrenGap: 10 }}>
          <Label>{editItemId ? '✏️ Edit Item' : '➕ Add New Item'}</Label>

          <TextField
            label="Title"
            value={title}
            onChange={(e, val) => this._handleInputChange(e, val, 'title')}
          />
          <TextField
            label="Body"
            multiline
            value={body}
            onChange={(e, val) => this._handleInputChange(e, val, 'body')}
          />
          <TextField
            label="Letter"
            value={letter}
            onChange={(e, val) => this._handleInputChange(e, val, 'letter')}
          />

          <PrimaryButton text={editItemId ? 'Update' : 'Create'} onClick={this._createOrUpdateItem} />
          {editItemId && (
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ title: '', body: '', letter: '', editItemId: null })}
            />
          )}
        </Stack>

        <hr />

        <Label>📋 Items from "FAQ" List</Label>

        {loading ? (
          <Label>Loading...</Label>
        ) : (
          <>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Title</th>
                  <th>Body</th>
                  <th>Letter</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {pagedItems.map(item => (
                  <tr key={item.Id}>
                    <td>{item.Id}</td>
                    <td>{item.Title}</td>
                    <td>{item.Body}</td>
                    <td>{item.Letter}</td>
                    <td>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />
                        <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._deleteItem(item.Id)} />
                      </Stack>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            {/* Pagination Controls */}
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 10 } }}>
              {Array.from({ length: totalPages }, (_, i) => (
                <DefaultButton
                  key={i}
                  text={`${i + 1}`}
                  onClick={() => this._handlePageChange(i + 1)}
                  styles={{
                    root: {
                      backgroundColor: currentPage === i + 1 ? '#0078d4' : 'transparent',
                      color: currentPage === i + 1 ? 'white' : 'black',
                      border: '1px solid #ccc',
                    },
                  }}
                />
              ))}
            </Stack>
          </>
        )}
      </section>
    );
  }
}

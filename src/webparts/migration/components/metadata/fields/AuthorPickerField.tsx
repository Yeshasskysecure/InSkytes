import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getKmsUsers, IKmsUser } from '../../../services/KmsUsersService';

export interface IAuthorPickerFieldProps {
  context: WebPartContext;
  selectedAuthorUpns?: string[];
  onChange: (upns: string[], users?: IKmsUser[]) => void;
  placeholder?: string;
  className?: string;
  density?: 'default' | 'compact';
}

const MIN_SEARCH_LENGTH = 1;
const MAX_RESULTS = 10;

const normalizeIdentityValue = (value?: string): string => {
  const rawValue = String(value || '').trim();
  if (!rawValue) {
    return '';
  }

  return rawValue.indexOf('|') >= 0
    ? rawValue.split('|').pop() || rawValue
    : rawValue;
};

const toUserMap = (user: IKmsUser): Record<string, IKmsUser> => {
  const entries: Record<string, IKmsUser> = {};
  [user.upn, user.email].filter(Boolean).forEach((identityValue) => {
    entries[normalizeIdentityValue(identityValue).toLowerCase()] = user;
  });
  return entries;
};

const dropdownStyles: Record<string, React.CSSProperties> = {
  wrapper: {
    width: '100%',
    position: 'relative',
    fontFamily: "'DM Sans', Arial, sans-serif"
  },
  inputWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    padding: '8px 16px',
    border: '1px solid #8f939a',
    borderRadius: '7px',
    background: '#fff',
    minHeight: '40px',
    alignItems: 'center',
    boxSizing: 'border-box',
    width: '100%',
    height: 'auto',
    overflow: 'visible'
  },
  token: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    background: '#f3f4f6',
    color: '#1d1d1f',
    borderRadius: '6px',
    padding: '4px 10px',
    fontSize: '13px',
    fontWeight: 500,
    maxWidth: '100%',
    boxSizing: 'border-box',
    lineHeight: 1.35
  },
  tokenClose: {
    cursor: 'pointer',
    color: '#64748b',
    border: 'none',
    background: 'none',
    padding: 0,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '16px',
    height: '16px',
    borderRadius: '999px',
    fontSize: '14px',
    lineHeight: 1,
    flexShrink: 0
  },
  input: {
    flex: 1,
    minWidth: '120px',
    border: 'none',
    outline: 'none',
    background: 'transparent',
    padding: '0',
    fontSize: '17px',
    color: '#1d1d1f',
    fontWeight: 400
  },
  menu: {
    position: 'absolute',
    top: 'calc(100% + 6px)',
    left: 0,
    right: 0,
    width: '100%',
    maxHeight: '220px',
    overflowY: 'auto',
    background: '#ffffff',
    border: '1px solid rgba(15,33,51,0.12)',
    borderRadius: '8px',
    boxShadow: '0 8px 18px rgba(15, 23, 42, 0.06)',
    padding: '4px 0',
    zIndex: 1000
  },
  option: {
    display: 'block',
    width: '100%',
    textAlign: 'left',
    padding: '10px 12px',
    background: 'transparent',
    border: 0,
    cursor: 'pointer',
    color: '#1b3752',
    fontSize: '17px',
    fontWeight: 500
  },
  meta: {
    display: 'block',
    fontSize: '12px',
    color: '#6b7280'
  },
  hint: {
    padding: '10px 12px',
    fontSize: '12px',
    color: '#6b7280'
  },
  error: {
    marginTop: '6px',
    fontSize: '12px',
    color: '#b42318'
  },
  loading: {
    marginTop: '6px',
    fontSize: '12px',
    color: '#6b7280'
  }
};

export const AuthorPickerField: React.FC<IAuthorPickerFieldProps> = ({
  context,
  selectedAuthorUpns,
  onChange,
  placeholder,
  className,
  density = 'default'
}) => {
  const [allUsers, setAllUsers] = React.useState<IKmsUser[]>([]);
  const [filteredUsers, setFilteredUsers] = React.useState<IKmsUser[]>([]);
  const [selectedUserMap, setSelectedUserMap] = React.useState<Record<string, IKmsUser>>({});
  const [isSearching, setIsSearching] = React.useState(false);
  const [error, setError] = React.useState('');
  const [query, setQuery] = React.useState('');
  const [isOpen, setIsOpen] = React.useState(false);
  const blurTimerRef = React.useRef<number | null>(null);
  const debounceRef = React.useRef<number | null>(null);

  const contextUser = React.useMemo<IKmsUser | null>(() => {
    const pageUser = context.pageContext.user as any;
    const currentEmail = normalizeIdentityValue(pageUser?.email || pageUser?.loginName || '');
    const displayName = String(pageUser?.displayName || pageUser?.title || '').trim();

    if (!currentEmail || !displayName) {
      return null;
    }

    return {
      displayName,
      upn: currentEmail,
      email: currentEmail
    };
  }, [context]);

  React.useEffect(() => {
    return () => {
      if (blurTimerRef.current) {
        window.clearTimeout(blurTimerRef.current);
      }
      if (debounceRef.current) {
        window.clearTimeout(debounceRef.current);
      }
    };
  }, []);

  React.useEffect(() => {
    let isDisposed = false;

    const loadUsers = async (): Promise<void> => {
      setError('');

      try {
        const users = await getKmsUsers(context);

        if (isDisposed) {
          return;
        }

        setAllUsers(users);
        setSelectedUserMap((currentMap) => {
          const nextMap = { ...currentMap };
          users.forEach((user) => {
            Object.assign(nextMap, toUserMap(user));
          });
          if (contextUser) {
            Object.assign(nextMap, toUserMap(contextUser));
          }
          return nextMap;
        });
      } catch (loadError) {
        if (isDisposed) {
          return;
        }

        console.error('Author list load failed:', loadError);
        setError(loadError instanceof Error ? loadError.message : 'Author list load failed.');
      }
    };

    void loadUsers();

    return () => {
      isDisposed = true;
    };
  }, [context, contextUser]);

  React.useEffect(() => {
    const selectedUpns = (selectedAuthorUpns || [])
      .map(normalizeIdentityValue)
      .filter(Boolean);

    if (selectedUpns.length === 0) {
      return;
    }

    if (contextUser) {
      const contextUserMap = toUserMap(contextUser);
      const isContextUserMissing = Object.keys(contextUserMap).some((key) => !selectedUserMap[key]);

      if (isContextUserMissing) {
        setSelectedUserMap((currentMap) => ({
          ...currentMap,
          ...contextUserMap
        }));
      }
    }

    const unresolvedUpns = selectedUpns.filter((upn) => !selectedUserMap[upn.toLowerCase()]);
    if (unresolvedUpns.length === 0) {
      return;
    }

    let isDisposed = false;

    const resolveSelectedAuthors = async (): Promise<void> => {
      try {
        const graphClient = await context.msGraphClientFactory.getClient('3');
        const resolvedUsers: IKmsUser[] = [];

        for (const upn of unresolvedUpns) {
          try {
            const escapedUpn = upn.replace(/'/g, "''");
            const response = await graphClient
              .api(`/users?$filter=userPrincipalName eq '${escapedUpn}' or mail eq '${escapedUpn}'&$select=displayName,mail,userPrincipalName&$top=1`)
              .version('v1.0')
              .get();
            const user = (response?.value || [])[0];

            if (user?.displayName) {
              const resolvedUpn = user.userPrincipalName || user.mail || upn;
              resolvedUsers.push({
                displayName: user.displayName,
                upn: resolvedUpn,
                email: user.mail || resolvedUpn
              });
            }
          } catch (selectedUserError) {
            console.warn('Selected author display lookup failed:', selectedUserError);
          }
        }

        if (isDisposed || resolvedUsers.length === 0) {
          return;
        }

        setSelectedUserMap((currentMap) => {
          const nextMap = { ...currentMap };
          resolvedUsers.forEach((user) => Object.assign(nextMap, toUserMap(user)));
          return nextMap;
        });
      } catch (loadError) {
        console.warn('Selected author resolver failed:', loadError);
      }
    };

    void resolveSelectedAuthors();

    return () => {
      isDisposed = true;
    };
  }, [context, contextUser, selectedAuthorUpns, selectedUserMap]);

  const selectedUsers = React.useMemo(() => {
    const selectedUpns = (selectedAuthorUpns || [])
      .map((upn) => (upn || '').trim())
      .filter(Boolean);

    return selectedUpns
      .map((upn) => {
        const normalizedUpn = normalizeIdentityValue(upn);
        const mappedUser = selectedUserMap[normalizedUpn.toLowerCase()];
        if (mappedUser) {
          return mappedUser;
        }

        if (contextUser && normalizeIdentityValue(contextUser.email).toLowerCase() === normalizedUpn.toLowerCase()) {
          return contextUser;
        }

        return {
          displayName: 'Loading author...',
          upn: normalizedUpn,
          email: normalizedUpn
        };
      });
  }, [contextUser, selectedAuthorUpns, selectedUserMap]);

  React.useEffect(() => {
    if (!isOpen) {
      setQuery('');
    }
  }, [isOpen]);

  const searchUsers = React.useCallback(async (searchQuery: string): Promise<void> => {
    const trimmedQuery = searchQuery.trim();
    if (!trimmedQuery || trimmedQuery.length < MIN_SEARCH_LENGTH) {
      setFilteredUsers([]);
      return;
    }

    setIsSearching(true);
    setError('');

    try {
      const selectedUpnLookup = new Set((selectedAuthorUpns || []).map((upn) => (upn || '').toLowerCase()));
      const normalizedQuery = trimmedQuery.toLowerCase();
      const sourceUsers = allUsers.length > 0 ? allUsers : await getKmsUsers(context);
      const users = sourceUsers
        .filter((user: IKmsUser) =>
          user.upn &&
          (
            user.displayName.toLowerCase().indexOf(normalizedQuery) >= 0 ||
            user.email.toLowerCase().indexOf(normalizedQuery) >= 0 ||
            user.upn.toLowerCase().indexOf(normalizedQuery) >= 0
          ) &&
          selectedUpnLookup.has(user.upn.toLowerCase()) === false
        )
        .slice(0, MAX_RESULTS);

      setFilteredUsers(users);
    } catch (loadError) {
      console.error('Author search failed:', loadError);
      setError(loadError instanceof Error ? loadError.message : 'Author search failed.');
      setFilteredUsers([]);
    } finally {
      setIsSearching(false);
    }
  }, [allUsers, context, selectedAuthorUpns]);

  const shouldShowResultsMenu = !error && isOpen && query.trim().length >= MIN_SEARCH_LENGTH;

  return (
    <div style={dropdownStyles.wrapper}>
      <div
        className={className}
        style={{
          ...dropdownStyles.inputWrapper,
          backgroundImage: 'none',
          paddingRight: '10px'
        }}
        onClick={() => {
        if (blurTimerRef.current) {
          window.clearTimeout(blurTimerRef.current);
          blurTimerRef.current = null;
        }
        setIsOpen(true);
      }}
      >
        {selectedUsers.map(user => (
          <div
            key={user.upn}
            style={{
              ...dropdownStyles.token,
              fontSize: density === 'compact' ? '13px' : dropdownStyles.token.fontSize
            }}
          >
            <span
              style={{
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                whiteSpace: 'normal',
                overflowWrap: 'anywhere',
                minWidth: 0
              }}
            >
              {user.displayName}
            </span>
            <button
              type="button"
              style={dropdownStyles.tokenClose}
              onMouseDown={(event) => event.preventDefault()}
              onClick={(e) => {
                e.stopPropagation();
                const removableValues = new Set([user.upn, user.email].filter(Boolean).map((value) => value.toLowerCase()));
                const nextUpns = (selectedAuthorUpns || []).filter(u => removableValues.has((u || '').toLowerCase()) === false);
                const nextUsers = selectedUsers.filter(u => removableValues.has((u.upn || '').toLowerCase()) === false);
                onChange(nextUpns, nextUsers);
              }}
            >
              ×
            </button>
          </div>
        ))}
        <input
          type="text"
          value={query}
          placeholder={placeholder || 'Search and add authors'}
          style={{
            ...dropdownStyles.input,
            backgroundImage: 'none',
            fontSize: density === 'compact' ? '13.6px' : dropdownStyles.input.fontSize
          }}
          autoComplete="off"
          onFocus={() => {
            if (blurTimerRef.current) {
              window.clearTimeout(blurTimerRef.current);
              blurTimerRef.current = null;
            }
            setIsOpen(true);
          }}
          onChange={(event) => {
            const value = event.target.value;
            setQuery(value);
            setIsOpen(true);

            if (debounceRef.current) {
              window.clearTimeout(debounceRef.current);
            }

            debounceRef.current = window.setTimeout(() => {
              void searchUsers(value);
            }, 300);
          }}
          onBlur={() => {
            blurTimerRef.current = window.setTimeout(() => {
              setIsOpen(false);
            }, 150);
          }}
        />
      </div>

      {error && <div style={dropdownStyles.error}>{error}</div>}

      {shouldShowResultsMenu && (
        <div style={dropdownStyles.menu}>
          {isSearching && <div style={dropdownStyles.hint}>Searching...</div>}

          {!isSearching && filteredUsers.length === 0 && query.trim().length >= MIN_SEARCH_LENGTH && (
            <div style={dropdownStyles.hint}>No users found</div>
          )}

          {!isSearching && filteredUsers.map((user) => (
            <button
              key={user.upn}
              type="button"
              style={{
                ...dropdownStyles.option,
                fontSize: density === 'compact' ? '13.6px' : dropdownStyles.option.fontSize
              }}
              onMouseDown={(event) => {
                event.preventDefault();
                const nextUsers = [...selectedUsers, user];
                setSelectedUserMap((current) => ({
                  ...current,
                  [user.upn.toLowerCase()]: user
                }));
                onChange(nextUsers.map((selectedUser) => selectedUser.upn), nextUsers);
                setQuery('');
                setFilteredUsers([]);
                setIsOpen(true);
              }}
            >
              {user.displayName}
              <span style={dropdownStyles.meta}>{user.email}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

export default AuthorPickerField;


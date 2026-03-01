import logging

from dateutil.parser import parse
from .utils import Pagination, NEXT_LINK_KEYWORD, ApiComponent
from .directory import User

log = logging.getLogger(__name__)

class Group(ApiComponent):
    """ A Microsoft O365 group """

    _endpoints = {
            'get_group_owners': '/groups/{group_id}/owners',
            'get_group_members': '/groups/{group_id}/members',
            'get_group_events': '/groups/{group_id}/calendar/events',
            'get_group_events_view': '/groups/{group_id}/calendar/calendarView',
    }

    member_constructor = User

    def __init__(self, *, parent=None, con=None, **kwargs):
        """ A Microsoft O365 group

        :param parent: parent object
        :type parent: Teams
        :param Connection con: connection to use if no parent specified
        :param Protocol protocol: protocol to use if no parent specified
         (kwargs)
        :param str main_resource: use this resource instead of parent resource
         (kwargs)
        """
        if parent and con:
            raise ValueError('Need a parent or a connection but not both')
        self.con = parent.con if parent else con

        cloud_data = kwargs.get(self._cloud_data_key, {})

        self.object_id = cloud_data.get('id')

        # Choose the main_resource passed in kwargs over parent main_resource
        main_resource = kwargs.pop('main_resource', None) or (
            getattr(parent, 'main_resource', None) if parent else None)

        main_resource = '{}{}'.format(main_resource, '')

        super().__init__(
            protocol=parent.protocol if parent else kwargs.get('protocol'),
            main_resource=main_resource)

        self.type = cloud_data.get('@odata.type')
        self.display_name = cloud_data.get(self._cc('displayName'), '')
        self.description = cloud_data.get(self._cc('description'), '')
        self.mail = cloud_data.get(self._cc('mail'), '')
        self.mail_nickname = cloud_data.get(self._cc('mailNickname'), '')
        self.visibility = cloud_data.get(self._cc('visibility'), '')

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return 'Group: {}'.format(self.display_name)

    def __eq__(self, other):
        return self.object_id == other.object_id

    def __hash__(self):
        return self.object_id.__hash__()

    def get_group_members(self, recursive=False):
        """ Returns members of given group
        :param bool recursive: drill down to users if group has other group as a member
        :rtype: list[User]
        """
        if recursive:
            recursive_data = self._get_group_members_raw()
            for member in recursive_data:
                if member['@odata.type'] == '#microsoft.graph.group':
                    recursive_members = Groups(con=self.con, protocol=self.protocol).get_group_by_id(member['id'])._get_group_members_raw()
                    recursive_data.extend(recursive_members)
            return [self.member_constructor(parent=self, **{self._cloud_data_key: lst}) for lst in recursive_data]
        else:
            return [self.member_constructor(parent=self, **{self._cloud_data_key: lst}) for lst in self._get_group_members_raw()]

    def _get_group_members_raw(self):
        url = self.build_url(self._endpoints.get('get_group_members').format(group_id=self.object_id))

        response = self.con.get(url)
        if not response:
            return []

        data = response.json()
        return data.get('value', [])

    def get_group_owners(self):
        """ Returns owners of given group

        :rtype: list[User]
        """
        url = self.build_url(self._endpoints.get('get_group_owners').format(group_id=self.object_id))

        response = self.con.get(url)
        if not response:
            return []

        data = response.json()

        return [self.member_constructor(parent=self, **{self._cloud_data_key: lst}) for lst in data.get('value', [])]


    def get_group_events(self, limit=25, *, query=None, order_by=None, batch=None,
                   download_attachments=False, include_recurring=True):
        """ Get events from the this Calendar

        :param int limit: max no. of events to get. Over 999 uses batch.
        :param query: applies a OData filter to the request
        :type query: Query or str
        :param order_by: orders the result set based on this condition
        :type order_by: Query or str
        :param int batch: batch size, retrieves items in
         batches allowing to retrieve more items than the limit.
        :param download_attachments: downloads event attachments
        :param bool include_recurring: whether to include recurring events or not
        :return: list of events in this calendar
        :rtype: list[Event] or Pagination
        """

        if include_recurring:
            url = self.build_url(
                self._endpoints.get('get_group_events_view').format(group_id=self.object_id))
        else:
            url = self.build_url(
                self._endpoints.get('get_group_events').format(group_id=self.object_id))

        if limit is None or limit > self.protocol.max_top_value:
            batch = self.protocol.max_top_value

        if batch:
            download_attachments = False

        params = {'$top': batch if batch else limit}

        if include_recurring:
            start = None
            end = None
            if query and not isinstance(query, str):
                # extract start and end from query because
                # those are required by a calendarView
                for query_data in query._filters:
                    if not isinstance(query_data, list):
                        continue
                    attribute = query_data[0]
                    # the 2nd position contains the filter data
                    # and the 3rd position in filter_data contains the value
                    word = query_data[2][3]

                    if attribute.lower().startswith('start/'):
                        start = word.replace("'", '')  # remove the quotes
                        query.remove_filter('start')
                    if attribute.lower().startswith('end/'):
                        end = word.replace("'", '')  # remove the quotes
                        query.remove_filter('end')

            if start is None or end is None:
                raise ValueError("When 'include_recurring' is True you must provide a 'start' and 'end' datetimes inside a Query instance.")

            if end < start:
                raise ValueError('When using "include_recurring=True", the date asigned to the "end" datetime'
                                 ' should be greater or equal than the date asigned to the "start" datetime.')

            params[self._cc('startDateTime')] = start
            params[self._cc('endDateTime')] = end

        if order_by:
            params['$orderby'] = order_by

        if query:
            if isinstance(query, str):
                params['$filter'] = query
            else:
                params.update(query.as_params())

        response = self.con.get(url, params=params,
                                headers={'Prefer': 'outlook.timezone="UTC"'})
        if not response:
            return iter(())

        data = response.json()

        # Everything received from cloud must be passed as self._cloud_data_key
        events = (self.event_constructor(parent=self,
                                         download_attachments=
                                         download_attachments,
                                         **{self._cloud_data_key: event})
                  for event in data.get('value', []))
        next_link = data.get(NEXT_LINK_KEYWORD, None)
        if batch and next_link:
            return Pagination(parent=self, data=events,
                              constructor=self.event_constructor,
                              next_link=next_link, limit=limit)
        else:
            return events

class Groups(ApiComponent):
    """ A microsoft groups class
        In order to use the API following permissions are required.
        Delegated (work or school account) - Group.Read.All, Group.ReadWrite.All
    """

    _endpoints = {
        'get_user_groups': '/users/{user_id}/memberOf',
        'get_group_by_id': '/groups/{group_id}',
        'get_group_by_mail': '/groups/?$search="mail:{group_mail}"&$count=true',
        'list_groups': '/groups',
    }

    group_constructor = Group

    def __init__(self, *, parent=None, con=None, **kwargs):
        """ A Teams object

        :param parent: parent object
        :type parent: Account
        :param Connection con: connection to use if no parent specified
        :param Protocol protocol: protocol to use if no parent specified
         (kwargs)
        :param str main_resource: use this resource instead of parent resource
         (kwargs)
        """
        if parent and con:
            raise ValueError('Need a parent or a connection but not both')
        self.con = parent.con if parent else con

        # Choose the main_resource passed in kwargs over the host_name
        main_resource = kwargs.pop('main_resource',
                                   '')  # defaults to blank resource
        super().__init__(
            protocol=parent.protocol if parent else kwargs.get('protocol'),
            main_resource=main_resource)

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return 'Microsoft O365 Group parent class'

    def get_group_by_id(self, group_id = None):
        """ Returns Microsoft O365/AD group with given id

        :param group_id: group id of group

        :rtype: Group
        """

        if not group_id:
            raise RuntimeError('Provide the group_id')

        if group_id:
            # get channels by the team id
            url = self.build_url(
                self._endpoints.get('get_group_by_id').format(group_id=group_id))

        response = self.con.get(url)

        if not response:
            return None

        data = response.json()

        return self.group_constructor(parent=self,
                                **{self._cloud_data_key: data})

    def get_group_by_mail(self, group_mail = None):
        """ Returns Microsoft O365/AD group by mail field

        :param group_name: mail of group

        :rtype: Group
        """
        if not group_mail:
            raise RuntimeError('Provide the group mail')

        if group_mail:
            # get groups by filter mail
            url = self.build_url(
                self._endpoints.get('get_group_by_mail').format(group_mail=group_mail))

        response = self.con.get(url, headers={'ConsistencyLevel': 'eventual'})

        if not response:
            return None

        data = response.json()

        if '@odata.count' in data and data['@odata.count'] < 1:
            raise RuntimeError('Not found group with provided filters')

        # mail is unique field so, we expect exact match -> always use first element from list
        return self.group_constructor(parent=self,
                                **{self._cloud_data_key: data.get('value')[0]})

    def get_user_groups(self, user_id = None):
        """ Returns list of groups that given user has membership

        :param user_id: user_id

        :rtype: list[Group]
        """

        if not user_id:
            raise RuntimeError('Provide the user_id')

        if user_id:
            # get channels by the team id
            url = self.build_url(
                self._endpoints.get('get_user_groups').format(user_id=user_id))

        response = self.con.get(url)

        if not response:
            return None

        data = response.json()

        return [
            self.group_constructor(parent=self, **{self._cloud_data_key: group})
            for group in data.get('value', [])]

    def list_groups(self):
        """ Returns list of groups
        :rtype: list[Group]
        """

        url = self.build_url(
            self._endpoints.get('list_groups'))

        response = self.con.get(url)

        if not response:
            return None

        data = response.json()

        return [
            self.group_constructor(parent=self, **{self._cloud_data_key: group})
            for group in data.get('value', [])]

//ckeyll.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CKeyedLinkedList object.

#ifndef __CKEYLL_H__
#define __CKEYLL_H__

#include <windows.h>
#include <wtypes.h>
#include "resource.h"

// forward referencing
class CKeyedLinkedListItem;

#define LIST_ERROR -1L

// one way Linked list class
class CKeyedLinkedList
{
	// mutual friends
	friend class CKeyedLinkedListItem;

private:
	// Private member variables
	CKeyedLinkedListItem * m_pRootItem;		// pointer to root of list
	CKeyedLinkedListItem * m_pLastItem;		// pointer to last item added (for speed of add)
	long m_lItemCount;						// count of items in list
	int m_LLError;
	long m_oid;

	// Private member methods
	void SetError(int LLError);				// internal routine to set error trap etc

public:
	// Public member variables
	enum
	{
		LLE_NOBOTHBEFOREAFTER	= -1,
		LLE_OUTOFMEMORY			= -2
	} LLErrors;

	// Public member methods
	CKeyedLinkedList();
	~CKeyedLinkedList();

	// delete a specific link given its index or key
	BOOL	Delete(BSTR bstrKey);

	// insert a link into the list, flags specify insertion details
	long	Insert(BSTR bstrKey, VARIANT * pvData);

	// returns next link in list given a specific link
	CKeyedLinkedListItem * GetNext(CKeyedLinkedListItem * pItem = NULL);

	// finds a link given it's index or key, returns item and previous item
	long	Find(BSTR bstrKey, CKeyedLinkedListItem **ppPrevItem, CKeyedLinkedListItem **ppItem);

	// "walk" or traverse 'n' links for a given starting position
	CKeyedLinkedListItem * TravsereN(CKeyedLinkedListItem * pStartItem, long lIndex, CKeyedLinkedListItem **ppPrevItem);

	// functions to retreive a specific link given its index or key
	BOOL	Get(long lIndex, CKeyedLinkedListItem **ppItem);

	// function to return no. of items in list
	long	GetCount();

	// function to retrieve last LL error
	int		GetError();

	// replace the data in a specific link given its index or key
	BOOL	Update(BSTR bstrKey, VARIANT * pvNewItem);

};

typedef CKeyedLinkedList *PCKeyedLinkedList;

class CKeyedLinkedListItem
	{
	// mutual friends
	friend class CKeyedLinkedList;

	protected:
		BSTR m_bstrKey;							// item key
		VARIANT m_vItemData;					// item data
		CKeyedLinkedListItem * m_pNextItem;		// pointer to next item in chain

	public:
		// parameterized constructor (no default constructor)
		CKeyedLinkedListItem(
			CKeyedLinkedListItem * pPrevItem = NULL,	// previous item in link, NULL = root item
			BSTR bstrKey = NULL,						// key for new item
			VARIANT * pvData = NULL);					// data for new item
	
		long m_oid;

		// default destructor
		~CKeyedLinkedListItem();

		// function to return key of current item
		BSTR GetKey();

		// function to return data of current item
		VARIANT * GetData();

		// function to return pointer to next item
		CKeyedLinkedListItem * GetNextItem();
	};

typedef CKeyedLinkedListItem *PCKeyedLinkedListItem;

#endif //#ifndef __CKEYLL_H__

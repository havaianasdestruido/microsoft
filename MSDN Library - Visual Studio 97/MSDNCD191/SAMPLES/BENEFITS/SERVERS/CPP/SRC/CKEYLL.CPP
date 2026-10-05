//ckeyll.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for CKeyedLinkedList object.

#pragma warning (disable : 4800)	// turns off bool warning

#include "support.h"
#include "ckeyll.h"
#include "cstrfunc.h"

CKeyedLinkedList::CKeyedLinkedList()
{
	m_pRootItem		= new CKeyedLinkedListItem();
	m_pLastItem		= m_pRootItem;
	m_lItemCount	= 0L;
}

CKeyedLinkedList::~CKeyedLinkedList()
{
	CKeyedLinkedListItem * pItem;
	CKeyedLinkedListItem * pNextItem;

	if (m_pRootItem)
	{
		pItem = m_pRootItem;

		while (pItem->GetNextItem())
		{
			pNextItem = pItem->GetNextItem();

			delete pItem;

			pItem = pNextItem;
		}

		delete pItem;
	}
}

//GetCount
//Purpose	- to return count of items in list
//Outputs	- (long) # of items in list
long CKeyedLinkedList::GetCount()
{
	return (m_lItemCount);
}

//SetError
//Purpose	- internal routine to set error trap, erros etc
//Inputs	- (LLError)    error flag
void CKeyedLinkedList::SetError(int LLError)
{
	m_LLError = LLError;
}

//SetError
//Purpose	- function to retrieve last LL error
//Outputs	- last known error
int CKeyedLinkedList::GetError()
{
	return (m_LLError);
}

//Insert
//Purpose	- to insert key and data into LL
//Inputs	- (bstrKey)		- key to data being inserted
//			- (pvData)		- data to be inserted into linked list
//Outputs	- (long) either number of items in list or error value
long CKeyedLinkedList::Insert(BSTR bstrKey, VARIANT * pvData)
{
	ASSERT(pvData != NULL);

	m_pLastItem = new CKeyedLinkedListItem(m_pLastItem, bstrKey, pvData);
	if (m_pLastItem == NULL)
	{
		SetError(LLE_OUTOFMEMORY);
		return (LIST_ERROR);
	}

	return (m_lItemCount++);
}

//TravsereN
//Purpose	- to "walk" or traverse 'n' links for a given starting position
//Inputs	- (pStartItem)	- node in linked list to start traversal
//			- (lIndex)		- number of nodes to traverse
//			- (ppPrevItem)	- pointer to previous node in list
//Outputs	- (CKeyedLinkedListItem *) pointer to retreived
CKeyedLinkedListItem * CKeyedLinkedList::TravsereN(CKeyedLinkedListItem * pStartItem, long lIndex, CKeyedLinkedListItem **ppPrevItem)
{
	CKeyedLinkedListItem * pItem;

	pItem = pStartItem;

	if (pStartItem == m_pRootItem)
		pItem = pItem->GetNextItem();

	while (lIndex--)
	{
		if (pItem == NULL)
			return (NULL);

		if (ppPrevItem)
			*ppPrevItem = pItem;

		pItem = pItem->GetNextItem();
	}

	return (pItem);
}

//Find
//Purpose	- to find a link given it's index or key, returns item and previous item
//Inputs	- (bstrKey)		- key to find
//			- (ppPrevItem)	- pointer to previous node in list
//			- (ppItem)		- pointer to found node in list
//Outputs	- (long)		- index to found node
long CKeyedLinkedList::Find(	BSTR bstrKey,
								CKeyedLinkedListItem **ppPrevItem,
								CKeyedLinkedListItem **ppItem)
{
	long lIndex;
	CKeyedLinkedListItem * pItem;
	CKeyedLinkedListItem * pPrevItem;

	lIndex		= 0;
	pItem		= m_pRootItem->m_pNextItem;
	pPrevItem	= (CKeyedLinkedListItem * )NULL;

	if (ppPrevItem)
		*ppPrevItem = (CKeyedLinkedListItem *)NULL;

	while (pItem)
	{
		if (CStrFnc::BSTRCmpi(bstrKey, pItem->m_bstrKey) == 0)
		{
			if (ppItem)
				*ppItem = pItem;
			if (ppPrevItem)
			{
				if (pPrevItem == NULL)
					*ppPrevItem = m_pRootItem;
				else
					*ppPrevItem = pPrevItem;
			}
			return (lIndex);
		}

		pPrevItem = pItem;
		pItem = pItem->m_pNextItem;
		lIndex++;
	}

	return (LIST_ERROR);
}

//Delete
//Purpose	- to delete a specific link given its index or key
//Inputs	- (bstrKey)		- key to delete
//Outputs	- (bool)		- boolean success flag
BOOL CKeyedLinkedList::Delete(BSTR bstrKey)
	{
	long lIndex;
	CKeyedLinkedListItem * pPrevItem;
	CKeyedLinkedListItem * pItem;

	lIndex = this->Find(bstrKey, &pPrevItem, &pItem);
	if (lIndex != LIST_ERROR)
	{
		pPrevItem->m_pNextItem = pItem->m_pNextItem;
		delete pItem;
		m_lItemCount--;
		return (TRUE);
	}
	return (FALSE);
}

//Update
//Purpose	- to replace the data in a specific link given its index or key
//Inputs	- (bstrKey)		- key to update
//			- (pvNewItem)	- data to be replaced into linked list
//Outputs	- (bool)		- boolean success flag
BOOL CKeyedLinkedList::Update(BSTR bstrKey, VARIANT * pvNewItem)
{
	CKeyedLinkedListItem * pItem;

	long lIndex = this->Find(bstrKey, NULL, &pItem);
	if (lIndex != LIST_ERROR)
	{
		::VariantClear(&pItem->m_vItemData);
		::VariantCopyInd(&pItem->m_vItemData, pvNewItem);
		return (TRUE);
	}
	return (FALSE);
}

//GetNext
//Purpose	- to return next link in list given a specific link
//Inputs	- (pItem)		- node to find next from
//Outputs	- (CKeyedLinkedListItem *)		- next node in list
CKeyedLinkedListItem * CKeyedLinkedList::GetNext(CKeyedLinkedListItem * pItem)
{
	return ((pItem == NULL) ? m_pRootItem : pItem->m_pNextItem);
}

//Get
//Purpose	- to retreive a specific link given its index or key
//Inputs	- (lIndex)		- index to node to get
//			- (ppItem)		- pointer to variable to hold return value
//Outputs	- (bool)		- boolean success flag
BOOL CKeyedLinkedList::Get(long lIndex, CKeyedLinkedListItem **ppItem)
{
	CKeyedLinkedListItem *pItem;

	pItem = this->TravsereN(m_pRootItem, lIndex, NULL);

	if (pItem)
	{
		*ppItem = pItem;
		return (TRUE);
	}

	return (FALSE);
}

// constructor/destructor
CKeyedLinkedListItem::CKeyedLinkedListItem(CKeyedLinkedListItem * pPrevItem, BSTR bstrKey, VARIANT * pvData)
{
	if (pPrevItem)
		pPrevItem->m_pNextItem = this;

	m_pNextItem = NULL;
	::VariantInit(&m_vItemData);
	m_bstrKey = NULL;
	if (bstrKey != NULL)
		m_bstrKey = ::SysAllocString(bstrKey);

	if (pvData != NULL)
		::VariantCopyInd(&m_vItemData, pvData);
}

CKeyedLinkedListItem::~CKeyedLinkedListItem()
{
	::VariantClear(&m_vItemData);
	if (m_bstrKey)
		::SysFreeString(m_bstrKey);
}

//GetKey
//Purpose	- to return key of current item
//Outputs	- (BSTR)		- key of node
BSTR CKeyedLinkedListItem::GetKey()
{
	return (m_bstrKey);
}

//GetData
//Purpose	- to return data of current item
//Outputs	- (VARIANT *)		- data of node
VARIANT * CKeyedLinkedListItem::GetData()
{
	return (&m_vItemData);
}

//GetNextIte
//Purpose	- to return pointer to next item
//Outputs	- (CKeyedLinkedListItem *)		- pointer to next item
CKeyedLinkedListItem * CKeyedLinkedListItem::GetNextItem()
{
	return (m_pNextItem);
}


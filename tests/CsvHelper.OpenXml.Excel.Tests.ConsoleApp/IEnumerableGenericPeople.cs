namespace CsvHelper.OpenXml.Excel.Tests.ConsoleApp;

using System;
using System.Collections;

public class IEnumerableGenericPeople : IEnumerable
{
    private GenericPerson[] GenericPeople;

    public IEnumerableGenericPeople(GenericPerson[] genericpeople)
    {
        GenericPeople = new GenericPerson[genericpeople.Length];

        for (int i = 0; i < genericpeople.Length; i++)
        {
            GenericPeople[i] = genericpeople[i];
        }
    }

    public IEnumerableGenericPeopleEnumerator GetEnumerator()
    {
        return new IEnumerableGenericPeopleEnumerator(GenericPeople);
    }

    IEnumerator IEnumerable.GetEnumerator() => GenericPeople.GetEnumerator();

}

public class IEnumerableGenericPeopleEnumerator : IEnumerator
{
    public GenericPerson[] GenericPeople;

    int Position = -1;

    object IEnumerator.Current { get => Current; }

    public GenericPerson Current
    {
        get
        {
            try
            {
                return GenericPeople[Position];
            }
            catch (IndexOutOfRangeException)
            {
                throw new InvalidOperationException();
            }
        }
    }

    public IEnumerableGenericPeopleEnumerator(GenericPerson[] list)
    {
        GenericPeople = list;
    }

    public bool MoveNext()
    {
        Position++;

        return Position < GenericPeople.Length;
    }

    public void Reset() => Position = -1;
}
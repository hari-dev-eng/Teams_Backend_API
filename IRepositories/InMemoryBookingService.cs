using Teams_Backend_API.Models.DTOs;
using System.Collections.Concurrent;
using Teams_Backend_API.Models;

public class InMemoryBookingService : IBookingService
{
    private readonly ConcurrentDictionary<Guid, Booking> _bookings = new();
    private readonly object _lock = new object();

    public Task<Booking> CreateBookingAsync(BookingDto dto, string userEmail, string userName)
    {
        if (dto.StartTime >= dto.EndTime)
            throw new ArgumentException("Start must be before end.");

        // Check for time conflicts for the same user
        var userBookings = _bookings.Values
            .Where(b => b.UserEmail.Equals(userEmail, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var conflict = userBookings.Any(b =>
            (dto.StartTime >= b.StartTime && dto.StartTime < b.EndTime) ||
            (dto.EndTime > b.StartTime && dto.EndTime <= b.EndTime) ||
            (dto.StartTime <= b.StartTime && dto.EndTime >= b.EndTime));

        if (conflict)
            throw new InvalidOperationException("Conflict with existing booking.");

        var booking = new Booking(
            Guid.NewGuid(),
            dto.Title,
            dto.StartTime,
            dto.EndTime,
            userEmail,
            userName ?? userEmail,
            dto.Location,
            DateTime.UtcNow
        );

        _bookings[booking.Id] = booking;
        return Task.FromResult(booking);
    }

    public Task<IReadOnlyList<Booking>> GetAllBookingsAsync()
    {
        var list = _bookings.Values
            .OrderBy(b => b.StartTime)
            .ToList()
            .AsReadOnly();

        return Task.FromResult((IReadOnlyList<Booking>)list);
    }

    public Task<IReadOnlyList<Booking>> GetBookingsForUserAsync(string userEmail)
    {
        var list = _bookings.Values
            .Where(b => b.UserEmail.Equals(userEmail, StringComparison.OrdinalIgnoreCase))
            .OrderBy(b => b.StartTime)
            .ToList()
            .AsReadOnly();

        return Task.FromResult((IReadOnlyList<Booking>)list);
    }
}